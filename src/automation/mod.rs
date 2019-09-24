#![allow(non_camel_case_types, non_snake_case, unused)]

mod interfaces;

use std::convert::{TryFrom, TryInto};
use std::fmt;

use winapi::shared::ntdef::NULL;
use winapi::um::combaseapi::{CoInitializeEx, CoUninitialize, CLSCTX_SERVER};
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
use rusty_winapi::smart_variant::SmartVariant;
use rusty_winapi::smart_idispatch::SmartIDispatch;

use super::global_context::V8GlobalContext;
use interfaces::*;

pub struct V8Application(AutoCOMInterface::<IDispatch>);

impl V8Application {
    pub fn new() -> Result<V8Application, HRESULT> {
        V8Application::new_from_class::<V83ApplicationClass>()
    }

    pub fn new_thin() -> Result<V8Application, HRESULT> {
        V8Application::new_from_class::<V83CApplicationClass>()
    }

    pub fn connect(&self, connection_string: &str) -> Result<V8GlobalContext, ()> {
        let mut conn1Cdb: LPDISPATCH = std::ptr::null_mut();

        match unsafe { self.0.call("Connect", &[connection_string.into()]) } {
            Ok(success) => match success {
                SmartVariant::Bool(success) => if success {
                        Ok(self.0.clone().try_into().unwrap())
                    } else { Err(()) }
                _ => Err(())
            }
            _ => Err(())                
        }
    }

    fn new_from_class<C>() -> Result<V8Application, HRESULT> where C: winapi::Class {
        let hr = unsafe { CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED) };
        if winerror::FAILED(hr) {
            return Err(hr);
        }

        Ok(V8Application(AutoCOMInterface::<IDispatch>::create_instance(
            &<C as Class>::uuidof(),
            std::ptr::null_mut(),
            CLSCTX_SERVER,
        )?))
    }
}

impl Drop for V8Application {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}
