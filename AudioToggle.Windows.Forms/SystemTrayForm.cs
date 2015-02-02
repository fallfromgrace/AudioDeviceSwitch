using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reactive;
using System.Reactive.Linq;
using System.Drawing;
using CoreAudio.Net;
using CoreAudio.Net.Internal;

namespace AudioToggle.Windows.Forms
{
    internal class MMDeviceInfo
    {
        public String FriendlyName
        {
            get { return this.GetDeviceProperty(PropertyKeys.Device.FriendlyName); }
        }

        public String FriendlyInterfaceName
        {
            get { return this.GetDeviceProperty(PropertyKeys.DeviceInterface.FriendlyName); }
        }

        public String Id
        {
            get
            {
                String id;
                Int32 hr = this.device.GetId(out id);
                return id;
            }
        }

        public DeviceState State
        {
            get
            {
                DeviceState state;
                Int32 hr = this.device.GetState(out state);
                return state;
            }
        }

        public MMDeviceInfo(IMMDevice device)
        {
            this.device = device;
            Int32 hr = this.device
                .OpenPropertyStore(StorageAccessMode.Read, out this.propertyStore);
        }


        private String GetDeviceProperty(PropertyKey key)
        {
            PropVariant variant;
            int hr = propertyStore.GetValue(ref key, out variant);
            String result;
            if (hr >= 0)
                result = variant.Value.ToString();
            else
                result = String.Empty;
            variant.Dispose();
            return result;
        }

        //private String 

        private readonly IMMDevice device;
        private readonly IPropertyStore propertyStore;
    }

    internal class SystemTrayForm : Form
    {
        public SystemTrayForm()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.

            base.OnLoad(e);
        }

        static IEnumerable<IMMDevice> GetDevices(
            DataFlow dataFlow, 
            DeviceState deviceStates)
        {
            MMDeviceEnumerator deviceEnumerator = new MMDeviceEnumerator();
            IMMDeviceCollection devices;
            Int32 hr = deviceEnumerator
                .EnumAudioEndpoints(dataFlow, deviceStates, out devices);
            Int32 deviceCount;
            hr = devices.GetCount(out deviceCount);
            for (Int32 deviceIndex = 0; deviceIndex < deviceCount; deviceIndex++)
            {
                IMMDevice device;
                hr = devices.Item(deviceIndex, out device);
                yield return device;
            }
        }

        static String GetDefaultDeviceId(
            DataFlow dataFlow, 
            Role role)
        {
            MMDeviceEnumerator deviceEnumerator = new MMDeviceEnumerator();
            IMMDevice defaultDevice;
            Int32 hr = deviceEnumerator
                .GetDefaultAudioEndpoint(dataFlow, role, out defaultDevice);
            MMDeviceInfo deviceInfo = new MMDeviceInfo(defaultDevice);
            return deviceInfo.Id;
        }

        static void SetDefaultDevice(String deviceId)
        {
            PolicyConfig config = new PolicyConfig();
            config.SetDefaultEndpoint(deviceId, Role.Console);
        }

        private void InitializeComponent()
        {
            this.contextMenu = new ContextMenu();
            this.notifyIcon = new NotifyIcon();
            this.notifyIcon.Text = "Audio Toggle";
            this.notifyIcon.Icon = Icon.FromHandle(Properties.Resources.TrayIcon.GetHicon());
            this.notifyIcon.ContextMenu = this.contextMenu;
            this.notifyIcon.Visible = true;

            Observable
                .FromEventPattern<EventHandler, EventArgs>(
                    handler => this.contextMenu.Popup += handler,
                    handler => this.contextMenu.Popup -= handler)
                .Select(_ => GetDevices(DataFlow.Render, DeviceState.All)
                    .Select(device => new MMDeviceInfo(device))
                    .Select(device => new 
                    {
                        FriendlyName = device.FriendlyName,
                        Id = device.Id,
                        IsDefault = device.Id == GetDefaultDeviceId(DataFlow.Render, Role.Multimedia)
                    }))
                .Subscribe(ds =>
                {
                    this.contextMenu.MenuItems.Clear();
                    MenuItem[] menuItems = ds
                        .Select(d => new MenuItem { Checked = d.IsDefault, Text = d.FriendlyName, Tag = d.Id })
                        .ToArray();
                    menuItems
                        .ToObservable()
                        .SelectMany(d => Observable
                            .FromEventPattern<EventHandler, EventArgs>(
                                handler => d.Click += handler,
                                handler => d.Click -= handler))
                        .Select(ep => ep.Sender as MenuItem)
                        .Select(mi => mi.Tag.ToString())
                        .Subscribe(id => SetDefaultDevice(id));
                    this.contextMenu.MenuItems.AddRange(menuItems);
                });
        }

        private IObservable<Unit> whenContextMenuPopup;
        private ContextMenu contextMenu;
        private NotifyIcon notifyIcon;
    }

}
