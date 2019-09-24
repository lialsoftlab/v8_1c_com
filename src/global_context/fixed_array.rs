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
use winapi::um::oaidl::{
    IDispatch, IDispatchVtbl, DISPID, DISPID_NEWENUM, DISPPARAMS, EXCEPINFO, LPDISPATCH, LPVARIANT, SAFEARRAY, VARIANT,
};
use winapi::um::oleauto::{SysStringLen, VariantClear, VariantInit, DISPATCH_METHOD};
use winapi::um::unknwnbase::{IUnknown, IUnknownVtbl};
use winapi::um::winnt::{HRESULT, LOCALE_USER_DEFAULT, LONG, LPCSTR, LPSTR, WCHAR};
use winapi::RIDL;

use rusty_winapi::auto_bstr::AutoBSTR;
use rusty_winapi::auto_com_interface::AutoCOMInterface;
use rusty_winapi::smart_idispatch::SmartIDispatch;
use rusty_winapi::smart_variant::SmartVariant;

use super::array::Array1C;
use super::V8GlobalContext;

#[derive(Clone)]
pub struct FixedArray1C(AutoCOMInterface<IDispatch>);

impl FixedArray1C {
    pub fn new_from(ctx: &mut V8GlobalContext, source: &Array1C) -> FixedArray1C {
        FixedArray1C(ctx.new_object("FixedArray", vec![source.clone().into()]).expect("1C fixed array"))
    }

    pub fn count(&self) -> u32 {
        let result = self.0.call("Count", &[]).unwrap();
        result.try_into().unwrap()
    }

    pub fn find(&self, value: SmartVariant) -> Option<u32> {
        match self.0.call("Find", &[value]).unwrap() {
            SmartVariant::Empty => None,
            x => Some(x.try_into().unwrap()),
        }
    }

    pub fn get(&self, index: u32) -> SmartVariant {
        self.0.call("Get", &[SmartVariant::from(index)]).unwrap()
    }

    pub fn ubound(&self) -> u32 {
        // TODO: тип Число может быть больше 32бит.
        self.0.call("UBound", &[]).unwrap().try_into().unwrap()
    }
}

impl TryFrom<SmartVariant> for FixedArray1C {
    type Error = ();

    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IDispatch(x) => {
                let result = x.get_ids_of_names(&["Ubound", "Set"], LOCALE_USER_DEFAULT);
                if SUCCEEDED(result.1) {
                    Ok(FixedArray1C(x))
                } else {
                    Err(())
                }
            }
            _ => Err(()),
        }
    }
}

impl From<FixedArray1C> for SmartVariant {
    #[inline]
    fn from(x: FixedArray1C) -> Self {
        x.0.into()
    }
}

pub struct FixedArray1CIterator(u32, FixedArray1C);
pub struct RefFixedArray1CIterator<'a>(u32, &'a FixedArray1C);

impl IntoIterator for FixedArray1C {
    type Item = SmartVariant;
    type IntoIter = FixedArray1CIterator;

    fn into_iter(self) -> Self::IntoIter {
        FixedArray1CIterator(0, self)
    }
}

impl iter::Iterator for FixedArray1CIterator {
    type Item = SmartVariant;

    fn next(&mut self) -> Option<Self::Item> {
        if self.0 < self.1.count() {
            let result = self.1.get(self.0);
            self.0 += 1;

            Some(result)
        } else {
            None
        }
    }
}

impl<'a> IntoIterator for &'a FixedArray1C {
    type Item = SmartVariant;
    type IntoIter = RefFixedArray1CIterator<'a>;

    fn into_iter(self) -> Self::IntoIter {
        RefFixedArray1CIterator(0, self)
    }
}

impl<'a> iter::Iterator for RefFixedArray1CIterator<'a> {
    type Item = SmartVariant;

    fn next(&mut self) -> Option<Self::Item> {
        if self.0 < self.1.count() {
            let result = self.1.get(self.0);
            self.0 += 1;

            Some(result)
        } else {
            None
        }
    }
}
