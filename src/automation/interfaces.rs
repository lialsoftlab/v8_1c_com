#![allow(non_camel_case_types, non_snake_case, unused)]

use std::any::Any;
use std::convert::{TryFrom, TryInto};
use std::error::Error;

use winapi::shared::guiddef::{IID_NULL, REFIID};
use winapi::shared::minwindef::{PUINT, UINT};
use winapi::shared::ntdef::{INT, PULONG, ULONG};
use winapi::shared::wtypes::{BSTR, DATE, VARIANT_BOOL};
use winapi::um::oaidl::{
    IDispatch, IDispatchVtbl, DISPID, DISPID_NEWENUM, DISPPARAMS, EXCEPINFO, LPDISPATCH, LPVARIANT, SAFEARRAY, VARIANT,
};
use winapi::um::oleauto::{SysStringLen, VariantClear, VariantInit, DISPATCH_METHOD};
use winapi::um::unknwnbase::{IUnknown, IUnknownVtbl};
use winapi::um::winnt::{HRESULT, LOCALE_USER_DEFAULT, LONG, LPCSTR, LPSTR, WCHAR};
use winapi::RIDL;

// 1C OLE Automation Server (1cv8.exe) class for 1C client app (V83.Application)
RIDL! {#[uuid(0xE92B75E3, 0x2EA1, 0x4FEC, 0xB4, 0x93, 0xCE, 0xF3, 0xEC, 0x59, 0xFC, 0xA6)]
class V83ApplicationClass;
}
pub type LPV83APPLICATION = *mut V83ApplicationClass;

// 1C OLE Automation Server (1cv8c.exe) class for 1C thin client app (V83C.Application)
RIDL! {#[uuid(0x27665EFC, 0x55CA, 0x4885, 0xB4, 0x91, 0x6B, 0x90, 0xF0, 0x4D, 0x62, 0x57)]
class V83CApplicationClass;
}
pub type LPV83CAPPLICATION = *mut V83CApplicationClass;
