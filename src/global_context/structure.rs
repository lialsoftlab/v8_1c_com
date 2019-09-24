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
use rusty_winapi::smart_variant::{AutoVariant, SmartVariant};

use super::fixed_structure::FixedStructure1C;
use super::V8GlobalContext;

#[derive(Clone)]
pub struct Structure1C(AutoCOMInterface<IDispatch>);

impl Structure1C {
    pub fn new(ctx: &mut V8GlobalContext) -> Structure1C {
        Structure1C(ctx.new_object("Structure", vec![]).expect("1C structure"))
    }

    pub fn new_from(ctx: &mut V8GlobalContext, source: &FixedStructure1C) -> Structure1C {
        Structure1C(ctx.new_object("Structure", vec![source.clone().into()]).expect("1C structure"))
    }

    pub fn new_by_keys_values(ctx: &mut V8GlobalContext, keys: &str, values: Vec<SmartVariant>) -> Structure1C {
        use std::iter;
        let params: Vec<SmartVariant> = iter::once(keys.into()).chain(values.into_iter()).collect();

        Structure1C(ctx.new_object("Structure", params).expect("1C structure"))
    }

    pub fn insert(&mut self, key: &str, value: SmartVariant) -> SmartVariant {
        self.0.call_mut("Insert", &[key.into(), value]).unwrap()
    }

    pub fn count(&self) -> u32 {
        self.0.call("Count", &[]).unwrap().try_into().unwrap()
    }

    pub fn clear(&mut self) -> SmartVariant {
         self.0.call_mut("Clear", &[]).unwrap()
    }

    pub fn property(&self, key: &str) -> Option<SmartVariant> {
        let mut v: VARIANT = AutoVariant::new().into();

        match self.0.call("Property", &[key.into(), (&mut v as LPVARIANT).into()]).unwrap() {
            SmartVariant::Bool(true) => Some(v.into()),
            _ => None,
        }
    }

    pub fn delete(&mut self, key: &str) -> SmartVariant {
        self.0.call_mut("Delete", &[key.into()]).unwrap()
    }

    pub fn get(&self, name: &str) -> SmartVariant {
        self.0.get(&name).unwrap()
    }
}

impl TryFrom<SmartVariant> for Structure1C {
    type Error = ();

    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IDispatch(x) => {
                let result = x.get_ids_of_names(&["Insert", "Count", "Clear", "Property"], LOCALE_USER_DEFAULT);
                if SUCCEEDED(result.1) {
                    Ok(Structure1C(x))
                } else {
                    Err(())
                }
            }
            _ => Err(()),
        }
    }
}

impl From<Structure1C> for SmartVariant {
    #[inline]
    fn from(x: Structure1C) -> Self {
        x.0.into()
    }
}

// TODO: Нужно реализовать IENumVariant
//pub struct Structure1CIterator(u32, Structure1C);
//pub struct RefStructure1CIterator<'a>(u32, &'a Structure1C);
//
//impl IntoIterator for Structure1C {
//    type Item = SmartVariant;
//    type IntoIter = Structure1CIterator;
//
//    fn into_iter(self) -> Self::IntoIter {
//        Structure1CIterator(0, self)
//    }
//}
//
//impl iter::Iterator for Structure1CIterator {
//    type Item = SmartVariant;
//
//    fn next(&mut self) -> Option<Self::Item> {
//        if self.0 < self.1.count() {
//            let result = self.1.get(self.0);
//            self.0 += 1;
//
//            Some(result)
//        } else {
//            None
//        }
//    }
//}
//
//impl<'a> IntoIterator for &'a Structure1C {
//    type Item = SmartVariant;
//    type IntoIter = RefStructure1CIterator<'a>;
//
//    fn into_iter(self) -> Self::IntoIter {
//        RefStructure1CIterator(0, self)
//    }
//}
//
//impl<'a> iter::Iterator for RefStructure1CIterator<'a> {
//    type Item = SmartVariant;
//
//    fn next(&mut self) -> Option<Self::Item> {
//        if self.0 < self.1.count() {
//            let result = self.1.get(self.0);
//            self.0 += 1;
//
//            Some(result)
//        } else {
//            None
//        }
//    }
//}
