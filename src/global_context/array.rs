#![allow(non_camel_case_types, non_snake_case, unused)]

use std::any::Any;
use std::convert::{TryFrom, TryInto};
use std::error::Error;
use std::iter;
use std::ops;

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

use super::fixed_array::FixedArray1C;
use super::V8GlobalContext;

#[derive(Clone)]
pub struct Array1C(AutoCOMInterface<IDispatch>);

impl Array1C {
    pub fn new(ctx: &mut V8GlobalContext) -> Array1C {
        Array1C(ctx.new_object("Array", vec![]).expect("1C array"))
    }

    pub fn with_dimensions(ctx: &mut V8GlobalContext, dimensions: &[u32]) -> Array1C {
        let dimensions: Vec<SmartVariant> = dimensions.iter().map(|&x| x.try_into().unwrap()).collect();

        Array1C(ctx.new_object("Array", dimensions).expect("1C array"))
    }

    pub fn from_fixed(ctx: &mut V8GlobalContext, source: &FixedArray1C) -> Array1C {
        let params: Vec<SmartVariant> = vec![source.clone().into()];

        Array1C(ctx.new_object("Array", params).expect("1C array"))
    }

    pub fn add(&mut self, value: SmartVariant) {
        self.0.call_mut("Add", &[value]).unwrap();
    }

    pub fn clear(&mut self) {
        self.0.call_mut("Clear", &[]).unwrap();
    }

    pub fn count(&self) -> u32 {
        let result = self.0.call("Count", &[]).unwrap();
        result.try_into().unwrap()
    }

    pub fn delete(&mut self, index: u32) -> SmartVariant {
        self.0.call_mut("Delete", &[SmartVariant::from(index)]).unwrap()
    }

    pub fn find(&self, value: SmartVariant) -> Option<i32> {
        match self.0.call("Find", &[value]).unwrap() {
            SmartVariant::Empty => None,
            x => Some(x.try_into().unwrap()),
        }
    }

    pub fn get(&self, index: u32) -> SmartVariant {
        self.0.call("Get", &[SmartVariant::from(index)]).unwrap()
    }

    pub fn insert(&mut self, index: u32, value: SmartVariant) -> SmartVariant {
        self.0.call_mut("Insert", &[SmartVariant::from(index), value]).unwrap()
    }

    pub fn set(&mut self, index: u32, value: SmartVariant) -> SmartVariant {
        self.0.call_mut("Set", &[SmartVariant::from(index), value]).unwrap()
    }

    pub fn ubound(&mut self) -> u32 {
        self.0.call("UBound", &[]).unwrap().try_into().unwrap()
    }
}

impl TryFrom<SmartVariant> for Array1C {
    type Error = ();

    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IDispatch(x) => {
                let result = x.get_ids_of_names(&["Ubound", "Set"], LOCALE_USER_DEFAULT);
                if SUCCEEDED(result.1) {
                    Ok(Array1C(x))
                } else {
                    Err(())
                }
            }
            _ => Err(()),
        }
    }
}

impl From<Array1C> for SmartVariant {
    #[inline]
    fn from(x: Array1C) -> Self {
        x.0.into()
    }
}

pub struct Array1CIterator(u32, Array1C);
pub struct RefArray1CIterator<'a>(u32, &'a Array1C);

impl IntoIterator for Array1C {
    type Item = SmartVariant;
    type IntoIter = Array1CIterator;

    fn into_iter(self) -> Self::IntoIter {
        Array1CIterator(0, self)
    }
}

impl iter::Iterator for Array1CIterator {
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

impl<'a> IntoIterator for &'a Array1C {
    type Item = SmartVariant;
    type IntoIter = RefArray1CIterator<'a>;

    fn into_iter(self) -> Self::IntoIter {
        RefArray1CIterator(0, self)
    }
}

impl<'a> iter::Iterator for RefArray1CIterator<'a> {
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
