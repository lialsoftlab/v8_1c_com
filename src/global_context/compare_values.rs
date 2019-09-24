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
pub struct CompareValues1C(AutoCOMInterface<IDispatch>);

impl CompareValues1C {
    pub fn new(ctx: &mut V8GlobalContext) -> CompareValues1C {
        CompareValues1C(ctx.new_object("CompareValues", vec![]).expect("1C compare values object"))
    }

    pub fn compare(&self, value1: SmartVariant, value2: SmartVariant) -> i32 {
        self.0.call("Compare", &[value1, value2]).unwrap().try_into().unwrap()
    }
}

impl TryFrom<SmartVariant> for CompareValues1C {
    type Error = ();

    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IDispatch(x) => {
                let result = x.get_ids_of_names(&["Compare"], LOCALE_USER_DEFAULT);
                if SUCCEEDED(result.1) {
                    Ok(CompareValues1C(x))
                } else {
                    Err(())
                }
            }
            _ => Err(()),
        }
    }
}

impl From<CompareValues1C> for SmartVariant {
    #[inline]
    fn from(x: CompareValues1C) -> Self {
        x.0.into()
    }
}
