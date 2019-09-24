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
use super::key_and_value::KeyAndValue1C;
use super::structure::Structure1C;

#[derive(Clone)]
pub struct FixedStructure1C(AutoCOMInterface<IDispatch>);

impl FixedStructure1C {
    pub fn new(ctx: &mut V8GlobalContext) -> FixedStructure1C {
        FixedStructure1C(ctx.new_object("FixedStructure", vec![]).expect("1C fixed structure"))
    }

    pub fn new_from(ctx: &mut V8GlobalContext, source: &Structure1C) -> FixedStructure1C {
        FixedStructure1C(ctx.new_object("FixedStructure", vec![source.clone().into()]).expect("1C fixed structure"))
    }

    pub fn new_by_keys_values(ctx: &mut V8GlobalContext, keys: &str, values: Vec<SmartVariant>) -> FixedStructure1C {
        use std::iter;
        let params: Vec<SmartVariant> = iter::once(keys.into()).chain(values.into_iter()).collect();

        FixedStructure1C(ctx.new_object("FixedStructure", params).expect("1C fixed structure"))
    }

    pub fn count(&self) -> u32 {
        let result = self.0.call("Count", &[]).unwrap();
        result.try_into().unwrap()
    }

    pub fn property(&self, key: &str) -> Option<SmartVariant> {
        let mut v: VARIANT = AutoVariant::new().into();

        match self.0.call("Property", &[key.into(), (&mut v as LPVARIANT).into()]).unwrap() {
            SmartVariant::Bool(true) => Some(v.into()),
            _ => None,
        }
    }

    pub fn get(&self, name: &str) -> SmartVariant {
        self.0.get(&name).unwrap()
    }
}

impl TryFrom<SmartVariant> for FixedStructure1C {
    type Error = ();

    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IDispatch(x) => {
                let result = x.get_ids_of_names(&["Count", "Property"], LOCALE_USER_DEFAULT);
                if SUCCEEDED(result.1) {
                    Ok(FixedStructure1C(x))
                } else {
                    Err(())
                }
            }
            _ => Err(()),
        }
    }
}

impl From<FixedStructure1C> for SmartVariant {
    #[inline]
    fn from(x: FixedStructure1C) -> Self {
        x.0.into()
    }
}

// TODO: Нужно реализовать IENumVariant
//pub struct FixedStructure1CIterator(u32, FixedStructure1C);
//pub struct RefFixedStructure1CIterator<'a>(u32, &'a FixedStructure1C);
//
//impl IntoIterator for FixedStructure1C {
//    type Item = SmartVariant;
//    type IntoIter = FixedStructure1CIterator;
//
//    fn into_iter(self) -> Self::IntoIter {
//        FixedStructure1CIterator(0, self)
//    }
//}
//
//impl iter::Iterator for FixedStructure1CIterator {
//    type Item = KeyAndValue1C;
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
//impl<'a> IntoIterator for &'a FixedStructure1C {
//    type Item = SmartVariant;
//    type IntoIter = RefFixedStructure1CIterator<'a>;
//
//    fn into_iter(self) -> Self::IntoIter {
//        RefFixedStructure1CIterator(0, self)
//    }
//}
//
//impl<'a> iter::Iterator for RefFixedStructure1CIterator<'a> {
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
