#![allow(non_camel_case_types, non_snake_case, unused)]

use std::any::Any;
use std::convert::{TryFrom, TryInto};
use std::error::Error;
use std::iter;

use winapi::shared::guiddef::{IID_NULL, REFIID};
use winapi::shared::minwindef::{PUINT, UINT};
use winapi::shared::ntdef::{INT, PULONG, ULONG};
use winapi::shared::winerror::*;
use winapi::shared::wtypes::*;
use winapi::um::oaidl::{IDispatch, IDispatchVtbl, DISPID, DISPID_NEWENUM, DISPPARAMS, EXCEPINFO, LPDISPATCH, LPVARIANT, SAFEARRAY, VARIANT, LPVARDESC};
use winapi::um::oleauto::{SysStringLen, VariantClear, VariantInit, DISPATCH_METHOD};
use winapi::um::unknwnbase::{IUnknown, IUnknownVtbl};
use winapi::um::winnt::{HRESULT, LOCALE_USER_DEFAULT, LONG, LPCSTR, LPSTR, WCHAR};
use winapi::RIDL;

use rusty_winapi::auto_bstr::AutoBSTR;
use rusty_winapi::auto_com_interface::AutoCOMInterface;
use rusty_winapi::smart_idispatch::SmartIDispatch;
use rusty_winapi::smart_variant::{AutoVariant, SmartVariant};

use super::V8GlobalContext;

#[derive(Clone)]
pub struct KeyAndValue1C(AutoCOMInterface<IDispatch>);

impl KeyAndValue1C {
    pub fn key(&self) -> SmartVariant {
        self.0.get("Key").unwrap()
    }

    pub fn value(&self) -> SmartVariant {
        self.0.get("Value").unwrap()
    }
}

impl TryFrom<SmartVariant> for KeyAndValue1C {
    type Error = ();

    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IDispatch(x) => {
                let result = x.get_ids_of_names(&["Key", "Value"], LOCALE_USER_DEFAULT);
                if SUCCEEDED(result.1) {
                    Ok(KeyAndValue1C(x))
                } else {
                    Err(())
                }
            }
            _ => Err(()),
        }
    }
}

impl From<KeyAndValue1C> for SmartVariant {
    #[inline]
    fn from(x: KeyAndValue1C) -> Self {
        x.0.into()
    }
}
