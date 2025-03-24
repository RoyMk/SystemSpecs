import wmi  # Requires: pip install wmi
import time
wmiQuery = wmi.WMI()
class ComputerSystem:
    def __init__(self):
        # Initialize attributes with default values
        self.Name = ""
        self.Domain = ""
        self.PartOfDomain = ""
        self.SystemType = ""
        self.PrimaryOwnerName = ""
        self.query_system()


    def get_all_properties(self):
        print("=" * 40)
        print("Hardware Information")
        print("=" * 40)
        all_prop = vars(self)
        for k, v in all_prop.items():
            print(f"{k}: {v}")
    def query_system(self):
        # Connect to WMI and query computer system information
        system_objects = wmiQuery.Win32_ComputerSystem()

        # WMI property names (case-sensitive)
        target_properties = ["Name", "Domain", "PartOfDomain", "SystemType", "PrimaryOwnerName"]

        for system_item in system_objects:
            # Set instance attributes dynamically
            for prop in target_properties:
                if hasattr(system_item, prop):
                    # Match WMI properties to class attributes (case-insensitive)
                    setattr(self, prop, getattr(system_item, prop))



class Win32Processor:
    def __init__(self):
        # Initialize attributes with default values
        self.DeviceID = ""
        self.Name = ""
        self.Caption = ""
        self.MaxClockSpeed = ""
        self.SocketDesignation = ""
        self.Manufacturer = ""
        self.query_system()


    def get_all_properties(self):
        print("\n")
        print("=" * 40)
        print("Processor Information")
        print("=" * 40)
        all_prop = vars(self)
        for k, v in all_prop.items():
            print(f"{k}: {v}")

    def query_system(self):
        system_objects = wmiQuery.Win32_Processor()
        target_properties = ["DeviceID", "Name", "Caption", "MaxClockSpeed", "SocketDesignation","Manufacturer"]

        for system_item in system_objects:
            for prop in target_properties:
                if hasattr(system_item, prop):
                    setattr(self, prop, getattr(system_item, prop))




class PhysicalMemory:
    def __init__(self):
        # Define properties for a single RAM stick
        self.Manufacturer = None
        self.PartNumber = None
        self.Capacity = None
        self.Speed = None
        self.DeviceLocator = None


class SystemMemory:
    def __init__(self):
        # A list to hold details of all RAM sticks
        self.PhysicalMemoryModules = []
        self.query_memory_explicit()
    def query_memory_explicit(self):
        system_objects = wmiQuery.Win32_PhysicalMemory()

        # Iterate through all RAM sticks returned by WMI query
        for system_item in system_objects:
            memory_module = PhysicalMemory()

            # Explicitly assign properties for each RAM stick
            memory_module.Manufacturer = system_item.Manufacturer
            memory_module.PartNumber = system_item.PartNumber
            memory_module.Capacity = system_item.Capacity
            memory_module.Speed = system_item.Speed
            memory_module.DeviceLocator = system_item.DeviceLocator

            # Add the populated memory module to the list
            self.PhysicalMemoryModules.append(memory_module)

    def get_all_properties(self):
        count = 0
        for module in memory_explicit.PhysicalMemoryModules:
            memory_sticks = vars(module)
            memory_sticks["Capacity"] = float(memory_sticks["Capacity"]) / 1_000_000_000
            memory_sticks["Speed"] = str(memory_sticks["Speed"]) + " MT/s"
            print("\n")
            print("=" * 40)
            print(f"System Memory Information [Ram Stick [{count}]]")
            print("=" * 40)
            count +=1
            for k, v in memory_sticks.items():

                print(f"{k}: {v}")


class OperatingSystem:
    def __init__(self):
        self.Caption = None
        self.BuildNumber = None
        self.LastBootUpTime = None
        self.OSArchitecture = None

        self.query()
    def query(self):
        system_objects = wmiQuery.Win32_OperatingSystem()

        # Iterate through all RAM sticks returned by WMI query
        for system_item in system_objects:
            # Explicitly assign properties for each RAM stick
            self.Caption = system_item.Caption
            self.BuildNumber = system_item.BuildNumber
            self.LastBootUpTime = system_item.LastBootUpTime
            self.OSArchitecture = system_item.OSArchitecture

    def get_all_properties(self):
        print("\n")
        print("=" * 40)
        print("Operating System Information")
        print("=" * 40)

        for k,v in vars(self).items():
            print(f"{k}: {v}")
        print("\n")


class StorageDevice:
    def __init__(self):
        self.Model = None
        self.Size = None
        self.MediaType = None
        self.FreeSpace = None


class StorageInfo:
    def __init__(self):
        self.StorageDevices = []
        self.query_storage()

    def query_storage(self):
        disk_drives = wmiQuery.Win32_DiskDrive()
        logical_disks = wmiQuery.Win32_LogicalDisk()

        for drive in disk_drives:
            device = StorageDevice()
            device.Model = drive.Model
            device.Size = drive.Size
            device.MediaType = drive.MediaType

            for disk in logical_disks:
                if disk.DeviceID == drive.DeviceID[0]:
                    device.FreeSpace = disk.FreeSpace
                    break

            self.StorageDevices.append(device)

    def get_all_properties(self):
        print("\n" + "=" * 40)
        print("Storage Information")
        print("=" * 40)
        for device in self.StorageDevices:
            for k, v in vars(device).items():
                if k in ["Size", "FreeSpace"]:
                    v = float(v) / (1024 ** 3) if v else None  # Convert to GB
                print(f"{k}: {v}")
            print("-" * 40)


class GPUInfo:
    def __init__(self):
        self.Name = None
        self.AdapterRAM = None
        self.query_gpu()

    def query_gpu(self):
        gpu = wmiQuery.Win32_VideoController()[0]  # Assuming primary GPU
        self.Name = gpu.Name
        self.AdapterRAM = gpu.AdapterRAM

    def get_all_properties(self):
        print("\n" + "=" * 40)
        print("GPU Information")
        print("=" * 40)
        for k, v in vars(self).items():
            if k == "AdapterRAM":
                v = float(v) / (1024**3) if v else None  # Convert to GB
            print(f"{k}: {v}")


class MotherboardInfo:
    def __init__(self):
        self.Manufacturer = None
        self.Product = None
        self.BIOSVersion = None
        self.query_motherboard()

    def query_motherboard(self):
        baseboard = wmiQuery.Win32_BaseBoard()[0]
        bios = wmiQuery.Win32_BIOS()[0]
        self.Manufacturer = baseboard.Manufacturer
        self.Product = baseboard.Product
        self.BIOSVersion = bios.Version

    def get_all_properties(self):
        print("\n" + "=" * 40)
        print("Motherboard Information")
        print("=" * 40)
        for k, v in vars(self).items():
            print(f"{k}: {v}")


class NetworkInfo:
    def __init__(self):
        self.Adapters = []
        self.query_network()

    def query_network(self):
        network_configs = wmiQuery.Win32_NetworkAdapterConfiguration(IPEnabled=True)
        for config in network_configs:
            self.Adapters.append({
                "Description": config.Description,
                "IPAddress": config.IPAddress,
                "MACAddress": config.MACAddress
            })

    def get_all_properties(self):
        print("\n" + "=" * 40)
        print("Network Information")
        print("=" * 40)
        for adapter in self.Adapters:
            for k, v in adapter.items():
                print(f"{k}: {v}")
            print("-" * 40)


class DisplayInfo:
    def __init__(self):
        self.Monitors = []
        self.Resolution = None
        self.query_display()

    def query_display(self):
        monitors = wmiQuery.Win32_DesktopMonitor()
        for monitor in monitors:
            self.Monitors.append(monitor.Name)

        video_controller = wmiQuery.Win32_VideoController()[0]  # Assuming primary display
        self.Resolution = f"{video_controller.CurrentHorizontalResolution}x{video_controller.CurrentVerticalResolution}"

    def get_all_properties(self):
        print("\n" + "=" * 40)
        print("Display Information")
        print("=" * 40)
        print(f"Monitors: {', '.join(self.Monitors)}")
        print(f"Resolution: {self.Resolution}")


memory_explicit = SystemMemory()
storage_info = StorageInfo()
gpu_info = GPUInfo()
motherboard_info = MotherboardInfo()
network_info = NetworkInfo()
display_info = DisplayInfo()



if __name__ == "__main__":
    memory_explicit = SystemMemory()
    processor = Win32Processor()
    csystem = ComputerSystem()
    operating_system = OperatingSystem()
    storage_info = StorageInfo()
    gpu_info = GPUInfo()
    motherboard_info = MotherboardInfo()
    network_info = NetworkInfo()

    operating_system.get_all_properties()
    csystem.get_all_properties()
    processor.get_all_properties()
    memory_explicit.get_all_properties()
    storage_info.get_all_properties()
    gpu_info.get_all_properties()
    motherboard_info.get_all_properties()
    network_info.get_all_properties()