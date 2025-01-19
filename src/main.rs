use std::process::{exit, Command};
use windows::Win32::Devices::Bluetooth::{
    BluetoothFindDeviceClose, BluetoothFindFirstDevice, BluetoothFindNextDevice,
    BLUETOOTH_DEVICE_INFO, BLUETOOTH_DEVICE_SEARCH_PARAMS,
};
use windows::Win32::Foundation::{FALSE, HANDLE, TRUE};

fn main() {
    unsafe {
        let device_search_params = BLUETOOTH_DEVICE_SEARCH_PARAMS {
            dwSize: std::mem::size_of::<BLUETOOTH_DEVICE_SEARCH_PARAMS>() as u32,
            fReturnAuthenticated: TRUE,
            fReturnRemembered: TRUE,
            fReturnUnknown: TRUE,
            fReturnConnected: TRUE,
            fIssueInquiry: FALSE,
            cTimeoutMultiplier: 5,
            hRadio: HANDLE::default(),
        };

        let mut device_info = BLUETOOTH_DEVICE_INFO {
            dwSize: std::mem::size_of::<BLUETOOTH_DEVICE_INFO>() as u32,
            ..Default::default()
        };

        let handle_to_bluetooth_device_find =
            BluetoothFindFirstDevice(&device_search_params, &mut device_info).unwrap();

        if handle_to_bluetooth_device_find.is_invalid() {
            println!("No bluetooth device found.");
            exit(0);
        }

        loop {
            println!(
                "Device name:          <{}>",
                String::from_utf16_lossy(&device_info.szName)
            );
            println!(
                "Address:              <{:X}>",
                device_info.Address.Anonymous.ullLong
            );
            println!("Connected:            <{}>", device_info.fConnected.as_bool());
            println!(
                "Remembered:           <{}>",
                device_info.fRemembered.as_bool()
            );
            println!(
                "Authenticated:        <{}>",
                device_info.fAuthenticated.as_bool()
            );
            println!(
                "COM Port:             <{}>",
                device_com_port(device_info).unwrap_or(String::from("NONE"))
            );
            println!("---------------------------------------------------------------------------");
            if BluetoothFindNextDevice(handle_to_bluetooth_device_find, &mut device_info).is_err() {
                break;
            }
        }
        BluetoothFindDeviceClose(handle_to_bluetooth_device_find).unwrap();
    }
}

unsafe fn device_com_port(device_info: BLUETOOTH_DEVICE_INFO) -> Option<String> {
    let bluetooth_device_address = format!("{:x}", device_info.Address.Anonymous.ullLong);
    let powershell_version = get_powershell_version()
        .expect("PowerShell not installed or unsupported version. (Supported version = 5 and 7");

    let mut cmd = Command::new("powershell");

    if powershell_version.as_str().eq("5") {
        cmd.arg("Get-WmiObject -Query \"SELECT DeviceID, PNPDeviceID FROM Win32_SerialPort\" | Where-Object { $_.PNPDeviceID -match \"BTHENUM\" } | Select-Object DeviceID, PNPDeviceID");
    } else if powershell_version.as_str().eq("7") {
        // Get-CimInstance -ClassName Win32_SerialPort | Where-Object { $_.PNPDeviceID -match "BTHENUM" } | Select-Object DeviceID, PNPDeviceID
        cmd.arg("Get-CimInstance -ClassName Win32_SerialPort | Where-Object { $_.PNPDeviceID -match \"BTHENUM\" } | Select-Object DeviceID, PNPDeviceID");
    } else {
        return None;
    }

    String::from_utf8(cmd.output().unwrap().stdout)
        .unwrap().lines()
        .filter(|line| line.starts_with("COM"))
        .filter(|line| line.contains(bluetooth_device_address.to_uppercase().as_str()))
        .nth(0)
        .map(|line| line.split_whitespace().nth(0))
        .flatten()
        .map(|x| x.to_string())
}

fn get_powershell_version() -> Option<String> {
    let output = Command::new("powershell")
        .arg("-Command")
        .arg("$PSVersionTable.PSVersion")
        .output()
        .unwrap();

    // Process the version string
    let version = String::from_utf8(output.stdout).unwrap()
        .lines()
        .filter(|line| line.chars().next().map(|c| c.is_numeric()).unwrap_or(false))
        .next()
        .unwrap_or("")
        .trim()
        .to_string();

    // Check if the version is valid (5 or 7)
    if version.starts_with('5') {
        Some(String::from("5"))
    } else if version.starts_with('7') {
        Some(String::from("7"))
    } else {
        None
    }
}
