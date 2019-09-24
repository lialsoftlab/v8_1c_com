#![allow(non_camel_case_types, non_snake_case, unused)]

pub mod array;
pub mod fixed_array;
pub mod fixed_structure;
pub mod key_and_value;
pub mod compare_values;
pub mod structure;
pub mod value_table;

use std::any::Any;
use std::convert::{TryFrom, TryInto};
use std::error::Error;

use winapi::shared::guiddef::{IID_NULL, REFIID};
use winapi::shared::minwindef::{PUINT, UINT};
use winapi::shared::ntdef::{INT, LCID, PULONG, ULONG};
use winapi::shared::winerror::*;
use winapi::shared::wtypes::{BSTR, DATE, VARIANT_BOOL};
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

pub struct V8GlobalContext(AutoCOMInterface<IDispatch>);

impl V8GlobalContext {
    pub fn new_object(&mut self, type_name: &str, params: Vec<SmartVariant>) -> Result<AutoCOMInterface<IDispatch>, String> {
        use std::iter;

        let params: Vec<SmartVariant> = iter::once(SmartVariant::from(type_name)).chain(params.into_iter()).collect();

        Ok(self.0.call("NewObject", &params).expect("1C ::NewObject()").try_into().unwrap())
    }

    pub fn string(&mut self, expression: SmartVariant) -> Result<String, String> {
        match self.0.call("String", &[expression]).expect("1C ::String()") {
            SmartVariant::Text(x) => Ok(x),
            _ => Err("Returned value isn't a string".into()),
        }
    }

    // pub fn type_of(&self, value: SmartVARIANT) -> SmartVARIANT
    // {
    //     self.call("XMLTypeOf", &[value]).unwrap()
    // }

    // pub fn type_(&self, type_name: &str) -> SmartVARIANT
    // {
    //     self.call("XMLType", &[SmartVARIANT::from(type_name)]).unwrap()
    // }

    // pub fn new_array(&self, dimensions: &[i32]) -> Array1C {
    //     Array1C::new(self, dimensions)
    // }

    // pub fn GetDBStorageStructureInfo(
    //     &self,
    //     metadata: &[&str],
    //     db_names: bool,
    // ) -> Result<V8KVTable, String> {
    //     let mut md_obj_names = self.new_array(&[metadata.len() as i32]);
    //     for &name in metadata {
    //         md_obj_names.add(SmartVARIANT::from(name).into());
    //     }

    //     V8KVTable::from_variant(
    //         self.call(
    //             "SetDBStorageStructureInfo",
    //             &[md_obj_names.as_variant(), SmartVARIANT::from(db_names)],
    //         )
    //         .expect("1C key-value table")
    //         .into(),
    //     )
    // }
}

impl TryFrom<AutoCOMInterface<IDispatch>> for V8GlobalContext {
    type Error = ();

    fn try_from(x: AutoCOMInterface<IDispatch>) -> Result<V8GlobalContext, Self::Error> {
        // check NewObject
        let result = x.get_ids_of_names(&["NewObject"], LOCALE_USER_DEFAULT);
        if SUCCEEDED(result.1) {
            Ok(V8GlobalContext(x))
        } else {
            Err(())
        }
    }
}
