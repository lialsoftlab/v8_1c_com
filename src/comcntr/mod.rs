#![allow(non_camel_case_types, non_snake_case, unused)]

mod interfaces;

use std::convert::{TryFrom, TryInto};
use std::fmt;

use winapi::shared::ntdef::NULL;
use winapi::um::combaseapi::{CoInitializeEx, CoUninitialize, CLSCTX_ALL};
use winapi::um::objbase::COINIT_MULTITHREADED;
use winapi::um::winuser;

use winapi::ctypes::c_void;
use winapi::shared::guiddef::{IID_NULL, REFIID};
use winapi::shared::minwindef::{BOOL, DWORD, PULONG, ULONG};
use winapi::shared::winerror;
use winapi::um::winnt::{HRESULT, LOCALE_USER_DEFAULT, LONG, LPCSTR, LPSTR, WCHAR};

use winapi::shared::guiddef::GUID;
use winapi::shared::minwindef::UINT;
use winapi::shared::winerror::*;
use winapi::shared::wtypes::{BSTR, VARIANT_BOOL, VT_DISPATCH, VT_UNKNOWN};
use winapi::um::oaidl::{IDispatch, IDispatchVtbl, DISPID, DISPID_NEWENUM, DISPPARAMS, EXCEPINFO, LPDISPATCH, LPVARIANT, VARIANT};
use winapi::um::oleauto::{SysStringLen, VariantClear, VariantInit, DISPATCH_METHOD};
use winapi::um::unknwnbase::{IUnknown, IUnknownVtbl, LPUNKNOWN};
use winapi::Class;
use winapi::{ENUM, RIDL, STRUCT};

use rusty_winapi::auto_bstr::AutoBSTR;
use rusty_winapi::auto_com_interface::AutoCOMInterface;

use super::global_context::V8GlobalContext;
use interfaces::*;

pub struct V8ComConnector(AutoCOMInterface<IV8COMConnector>);

impl V8ComConnector {
    pub fn new() -> Result<V8ComConnector, HRESULT> {
        let hr = unsafe { CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED) };
        if winerror::FAILED(hr) {
            return Err(hr);
        }

        Ok(V8ComConnector(AutoCOMInterface::<IV8COMConnector>::create_instance(
            &<V8COMConnectorClass as Class>::uuidof(),
            std::ptr::null_mut(),
            CLSCTX_ALL,
        )?))
    }

    pub fn connect(&self, connection_string: &str) -> Result<V8GlobalContext, ()> {
        let connection_string: AutoBSTR = connection_string.try_into().unwrap();

        let mut conn1Cdb: LPDISPATCH = std::ptr::null_mut();

        let hr = unsafe { self.0.as_inner().Connect(*connection_string, &mut conn1Cdb) };

        if SUCCEEDED(hr) {
            Ok(AutoCOMInterface::<IDispatch>::try_from(conn1Cdb).unwrap().try_into().unwrap())
        } else {
            Err(())
        }
    }
}

impl Drop for V8ComConnector {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}
