extern crate dirs;
extern crate winapi;

use semver::Version;
use std::env;
use std::ffi::OsStr;
use std::fs;
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::ptr::null_mut;
use winapi::shared::guiddef::GUID;
use winapi::shared::winerror;
use winapi::shared::wtypesbase::CLSCTX_INPROC_SERVER;
use winapi::um::combaseapi::{CoCreateInstance, CoInitializeEx, CoUninitialize};
use winapi::um::objbase::{COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE};
use winapi::um::objidl::IPersistFile;
use winapi::um::shobjidl_core::IShellLinkW;
use winapi::Interface;

const CLSID_SHELL_LINK: GUID = GUID {
    Data1: 0x00021401,
    Data2: 0x0000,
    Data3: 0x0000,
    Data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

fn main() {
    let discord_path = find_discord();

    if discord_path.is_none() {
        return;
    }

    let discord_path = discord_path.unwrap();

    let discord_exe_path = discord_path.join("Discord.exe");

    create_shortcut(
        &discord_path.to_str().unwrap().replace('"', ""),
        &discord_exe_path.to_str().unwrap().replace('"', ""),
    );
}

fn find_discord() -> Option<PathBuf> {
    let username = env::var("LOCALAPPDATA").expect("USERNAME not found");

    let path = format!("{}\\Discord", username);

    println!("Looking for Discord in {:?}", path);

    let discord_path = Path::new(&path);

    let version = find_newest_discord_version(discord_path);

    if let Some(version) = version {
        println!("Found Discord version: {:?}", version);
        Some(version)
    } else {
        println!("No Discord version found");
        None
    }
}

fn find_newest_discord_version(base_path: &Path) -> Option<PathBuf> {
    let folders = fs::read_dir(base_path).expect("Failed to read directory");

    // Filter and map directory entries to (PathBuf, Version)
    let mut versioned_folders: Vec<(PathBuf, Version)> = folders
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .filter(|path| {
            path.is_dir()
                && path
                    .file_name()
                    .unwrap()
                    .to_str()
                    .unwrap()
                    .starts_with("app-")
        })
        .filter_map(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .and_then(|name| {
                    name.strip_prefix("app-").and_then(|version| {
                        Version::parse(version).ok().map(|ver| (path.clone(), ver))
                    })
                })
        })
        .collect();

    // Sort folders by version, newest first
    versioned_folders.sort_by(|a, b| b.1.cmp(&a.1));

    // Return the path of the newest version, if any
    versioned_folders.into_iter().map(|(path, _)| path).next()
}

fn create_shortcut(discord_path: &String, discord_exe_path: &String) {
    println!(
        "Creating shortcut for Discord at {:?}",
        discord_exe_path.replace('"', "")
    );

    unsafe {
        // Initialize COM library
        let res = CoInitializeEx(
            null_mut(),
            COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE,
        );

        if (res != winerror::S_OK) && (res != winerror::S_FALSE) {
            panic!("Failed to initialize COM library");
        }

        let mut psl: *mut IShellLinkW = null_mut();
        // Ensure to use CLSID_ShellLink and IID_IShellLinkW for CoCreateInstance
        let hres = CoCreateInstance(
            &CLSID_SHELL_LINK,
            null_mut(),
            CLSCTX_INPROC_SERVER,
            &IShellLinkW::uuidof(),
            &mut psl as *mut *mut _ as *mut _,
        );
        if hres == winerror::S_OK {
            let psl = &mut *psl;

            // Set the path to the executable and other properties here...
            let target_path = OsStr::new(&discord_exe_path.replace('"', ""))
                .encode_wide()
                .chain(Some(0))
                .collect::<Vec<u16>>();
            let start_directory = OsStr::new(&discord_path.replace('"', ""))
                .encode_wide()
                .chain(Some(0))
                .collect::<Vec<u16>>();
            let arguments = OsStr::new("--multi-instance")
                .encode_wide()
                .chain(Some(0))
                .collect::<Vec<u16>>();
            psl.SetPath(target_path.as_ptr());
            psl.SetArguments(arguments.as_ptr());
            psl.SetWorkingDirectory(start_directory.as_ptr());

            // Query the IShellLink object for the IPersistFile interface
            let mut ppf: *mut IPersistFile = null_mut();
            psl.QueryInterface(&IPersistFile::uuidof(), &mut ppf as *mut *mut _ as *mut _);
            if !ppf.is_null() {
                let ppf = &*ppf;

                let desktop_path = dirs::desktop_dir().expect("Could not find desktop path");
                let shortcut_path = desktop_path.join("Discord多開.lnk");

                let shortcut_path_str = shortcut_path.to_str().unwrap();

                // Convert the shortcut path to a wide string
                let wstr_path: Vec<u16> = OsStr::new(shortcut_path_str)
                    .encode_wide()
                    .chain(Some(0))
                    .collect();

                // Save the shortcut
                ppf.Save(wstr_path.as_ptr(), 1);

                // Release the IPersistFile interface
                ppf.Release();
            }

            // Release the IShellLink interface
            psl.Release();
        }

        // Uninitialize COM
        CoUninitialize();
    }
}
