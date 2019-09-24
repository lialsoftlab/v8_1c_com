#![allow(non_camel_case_types, non_snake_case, unused)]

use std::any::Any;
use std::convert::{TryFrom, TryInto};
use std::error::Error;
use std::iter;
use std::ops;

use rusty_winapi::auto_bstr::AutoBSTR;
use rusty_winapi::auto_com_interface::AutoCOMInterface;
use rusty_winapi::smart_idispatch::SmartIDispatch;
use rusty_winapi::smart_variant::SmartVariant;
use winapi::RIDL;
use winapi::shared::guiddef::{IID_NULL, REFIID};
use winapi::shared::minwindef::{PUINT, UINT};
use winapi::shared::ntdef::{INT, PULONG, ULONG};
use winapi::shared::winerror::*;
use winapi::shared::wtypes::*;
use winapi::um::oaidl::{
    DISPID, DISPID_NEWENUM, DISPPARAMS, EXCEPINFO, IDispatch, IDispatchVtbl, LPDISPATCH, LPVARIANT, SAFEARRAY, VARIANT,
};
use winapi::um::oleauto::{DISPATCH_METHOD, SysStringLen, VariantClear, VariantInit};
use winapi::um::unknwnbase::{IUnknown, IUnknownVtbl};
use winapi::um::winnt::{HRESULT, LOCALE_USER_DEFAULT, LONG, LPCSTR, LPSTR, WCHAR};

use super::V8GlobalContext;
use super::array::Array1C;
use super::compare_values::CompareValues1C;
use super::structure::Structure1C;

/// ## A table of values (ТаблицаЗначений)
/// 
/// The object is intended for storage of values in the form of a table. All basic operations with table data can be made through
/// this object. This object manages the rows of the value table and provides access to the collection of columns. Values in the
/// columns of a table can be of various types (including multi-types).
/// 
/// ### Availability 
/// Available on a server, thick client, external connection, and a mobile app (server).
/// 
/// ### Exchange
/// Possible exchange with the application server. This object can be serialized into/from XDTO. The XDTO type name of this object 
/// is the ValueTable, defined in *http://v8.1c.ru/8.1/data/core* namespace.
#[derive(Clone)]
pub struct ValueTable(AutoCOMInterface<IDispatch>);


/// ## A row of the values table
/// 
/// The object is describing a single row of the values table. 
///
/// ### Availability 
/// Available on a server, thick client, external connection, and a mobile app (server).
#[derive(Clone)]
pub struct ValueTableRow(AutoCOMInterface<IDispatch>);

/// ## A column of the values table
/// 
/// The object is describing a single column of the values table. 
/// Object instances are accessible from properties of the values table columns collection.
///
/// ### Availability 
/// Available on a server, thick client, external connection, and a mobile app (server).
#[derive(Clone)]
pub struct ValueTableColumn(AutoCOMInterface<IDispatch>);

pub enum ValueTableColumnBy {
    Index(u32),
    Name(String),
    Column(ValueTableColumn),
}

pub enum ValueTableRowBy {
    Index(u32),
    Row(ValueTableRow),
}

impl ValueTable {
    /// Construct a new empty values table. 
    /// 
    /// *Новый ТаблицаЗначений;*
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn new(ctx: &mut V8GlobalContext) -> ValueTable {
        ValueTable(ctx.new_object("ValueTable", vec![]).expect("Expected 1C values table!"))
    }

    /// Insert and return new empty ValueTableRow in the values table at position specified in *index*. 
    /// 
    /// *ОбъектТЗ.Вставить(<Индекс>);*
    /// 
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn insert(&mut self, index: u32) -> ValueTableRow {
        self.0.call_mut("Insert", &[index.into()]).unwrap().try_into().expect("Expected 1C values table row!")
    }

    /// Show a modal dialog box in the 1C thick client for interactive selection of a row from the values table, then returned as 
    /// Some(ValueTableRow) if it was selected or None if operation cancelled by user.
    /// 
    /// *ОбъектТЗ.ВыбратьСтроку(<Заголовок>, <НачальнаяСтрока>);*
    /// 
    /// ### Parameters
    /// - *header*: an optional string to show in dialog title bar (may be used as hint to user).
    /// - *start_row*: an optional values table row that will be pre-selected on dialog opening.
    /// 
    /// Available on a thick client only.
    pub fn choose_row(&self, header: Option<&str>, start_row: Option<ValueTableRow>) -> Option<ValueTableRow> {
        self.0.call("ChooseRow", &[header.into(), start_row.into()]).unwrap().into_option()
            .expect("Expected optional 1C values table row!")
    }

    /// Unload values from the specified column of the values table, into new 1C array and return it.
    /// 
    /// *ОбъектТЗ.ВыгрузитьКолонку(<Колонка>);*
    /// 
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn unload_column(&mut self, column: ValueTableColumnBy) -> Array1C {
        self.0.call_mut("UnloadColumn", &[column.into()]).unwrap().try_into()
            .expect("Expected 1C array!")
    }

    /// Add new row at the and of the values table and return it.
    /// 
    /// *ОбъектТЗ.Добавить();* 
    /// 
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn add(&mut self) -> ValueTableRow {
        self.0.call_mut("Insert", &[]).unwrap().try_into().expect("Expected 1C values table row!")
    }

    /// Load values into the specified column of the values table, from a 1C array, values are loaded in order of indexes.
    /// 
    /// *ОбъектТЗ.ЗагрузитьКолонку(<Массив>, <Колонка>);*
    /// 
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn load_column(&mut self, array: Array1C, column: ValueTableColumnBy) -> SmartVariant {
        self.0.call_mut("LoadColumn", &[array.into(), column.into()]).unwrap()
    }

    /// Fill all rows of the values table with specified value.
    /// 
    /// *ОбъектТЗ.ЗаполнитьЗначения(<Значение>, <Колонки>);*
    /// 
    /// ### Parameters
    /// - *value*: the value to fill in.
    /// - *columns*: an optional string with a comma-separated list of column names to fill in, or else if it's empty or None all
    /// table will be filled with value.
    /// 
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn fill_values(&mut self, value: SmartVariant, columns: Option<&str>) -> SmartVariant {
        self.0.call_mut("FillValues", &[columns.into()]).unwrap()
    }

    /// Get the numeric index of specified row of the values table.
    /// 
    /// *ОбъектТЗ.Индекс(<Строка>);*
    /// 
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn index_of(&self, row: ValueTableRow) -> u32 {
        self.0.call("IndexOf", &[row.into()]).unwrap().try_into().expect("Expected unsigned integer in VARIANT")
    }

    /// Get totals for specified column of the values table
    /// 
    /// *ОбъектТЗ.Итог(<Колонка>);*
    /// 
    /// Returns sum of values in all rows of specified column.
    /// If the column having a type defined and it's only one, then an attempt will be made to convert the values into the Number 
    /// type. If the column doesn't have a type defined or having multi-type, and it has the Number type defined in a set, then 
    /// during summation, will be used only values, which has the Number type, all other values will be ignored. If a multi-type 
    /// type is defined for the column, and a set does not have the Number type, then, as a result, it gets the Undefined value.
    /// 
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn total(&self, column: &str) -> Option<f64> {
        self.0.call("Total", &[column.into()]).unwrap().into_option().expect("Expected optional float in VARIANT")
    }

    /// Get a count of rows in the values table.
    ///
    /// *ОбъектТЗ.Количество();*
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn count(&self) -> u32 {
        self.0.call("Count", &[]).unwrap().try_into().expect("Expected unsigned integer in VARIANT")
    }

    /// Find a row with value in the specified columns.
    ///
    /// Method is most effective for unique values.
    ///
    /// *ОбъектТЗ.Найти(<Значение>, <Колонки>);*
    ///
    /// ### Parameters
    /// - *value*: the value to find.
    /// - *columns*: an optional string with a comma-separated list of column names to scan, or else if it's empty or None all table
    /// will be scanned for value.
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn find(&self, value: SmartVariant, columns: Option<&str>) -> Option<ValueTableRow> {
        self.0.call("Find", &[columns.into()]).unwrap().into_option()
            .expect("Expected optional 1C values table row!")
    }

    /// Find a rows in the values table that match the specified search conditions.
    ///
    /// Method is most effective for non-unique values.
    ///
    /// If indexes added to the values table, the index for the search is selected by exact matches of sets of columns in the index
    /// and in search conditions, the order of columns doesn't matter.
    ///
    /// *ОбъектТЗ.НайтиСтроки(<ПараметрыОтбора>);*
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn find_rows(&self, filter: Structure1C) -> Array1C {
        self.0.call("FindRows", &[filter.into()]).unwrap().try_into()
            .expect("Expected 1C array!")
    }

    /// Remove all rows from the values table, columns structure stays unchanged.
    ///
    /// *ОбъектТЗ.Очистить();*
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn clear(&mut self) -> SmartVariant {
        self.0.call("Clear", &[]).unwrap()
    }

    /// Gets row from the values table by a numeric index.
    ///
    /// *ОбъектТЗ.Получить(<Индекс>);*
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn get(&self) -> ValueTableRow {
        self.0.call("Get", &[]).unwrap().try_into().expect("Expected 1C values table row!")
    }

    /// Roll-up specified columns in the values table.
    ///
    /// Rows with the same values in the columns specified in the *groupby_columns* parameter are combined into one, the values of
    /// these rows in the columns specified in the *sum_columns* parameter are accumulated.
    ///
    /// **Important!** Lists of columns should not overlap. Columns that are not included in any of the column lists are deleted
    /// from the value table after the method is executed.
    ///
    /// *ОбъектТЗ.Свернуть(<КолонкиГруппировок>, <КолонкиСуммирования>);*
    ///
    /// ### Parameters
    /// - *groupby_columns*: a string with a comma-separated list of columns names to group by.
    /// - *sum_columns*: an optional string with a comma-separated list of columns names to sum.
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn group_by(&mut self, groupby_columns: &str, sum_columns: Option<&str>) -> SmartVariant {
        self.0.call_mut("GroupBy", &[groupby_columns.into(), sum_columns.into()]).unwrap()
    }

    /// Move the row on the specified offset in the values table.
    ///
    /// Positive offset value -- move the row down (to the end), negative offset value -- move the row up (to the beginning).
    ///
    /// *ОбъектТЗ.Сдвинуть(<Строка>, <Смещение>);*
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn move_row(&mut self, row: ValueTableRowBy, offset: i32) -> SmartVariant {
        self.0.call_mut("Move", &[row.into(), offset.into()]).unwrap()
    }

    /// Copy specified columns with values from the specified rows of values table in a new values table.
    ///
    /// *ОбъектТЗ.Скопировать(<Строки>, <Колонки>);* or *ОбъектТЗ.Скопировать(<ПараметрыОтбора>, <Колонки>);*
    ///
    /// ### Parameters
    /// - *rows*: an optional array with the values table rows to copy or 1C structure with a rows filter conditions, or else if
    /// it's empty or None method will copy all rows.
    /// - *columns*: an optional string with a comma-separated list of column names to copy, or else if it's empty or None method
    /// will copy all columns.
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn copy(&self, rows: Option<SmartVariant>, columns: Option<&str>) -> ValueTable {
        self.0.call("Copy", &[rows.into(), columns.into()]).unwrap().try_into().expect("Expected 1C values table!")
    }

    /// Copy the specified columns (in a comma-separated list string) without values from the values table in a new values table.
    /// 
    /// *ОбъектТЗ.СкопироватьКолонки(<Колонки>);*
    /// 
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn copy_columns(&self, columns: Option<&str>) -> ValueTable {
        self.0.call("CopyColumns", &[columns.into()]).unwrap().try_into()
            .expect("Expected 1C values table!")
    }

    /// Sort the values table by specified rules.
    ///
    /// *ОбъектТЗ.Сортировать(<Колонки>, <ОбъектСравнения>);*
    ///
    /// ### Parameters
    /// - *columns*: a string with a comma-separated list of column names to sort by. After the name of a column with space can be
    /// specified sort ordering token. The sort order can be defined as "Desc" ("Убыв") to descending order and as "Asc" ("Возр") to
    /// ascending sort order. Sort order is determined by columns order in the string and if not specified directly by default used
    /// the ascending order for each column.
    /// - *sort_object*: an optional object for comparison of values of type CompareValues (СравнениеЗначений). Independent of a
    /// fact that such an object is specified in the method call, items with equal types are compared by type code, and items of
    /// simple types are compared by value. And additionally to this: if comparison object isn't specified than items of other types
    /// are compared by string representation. Else if comparison object is defined then objects compared by identifier, time
    /// moments compared by date and object identifier, and items of other types compared by a string representation.
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn sort(&mut self, columns: &str, sort_object: CompareValues1C) -> SmartVariant {
        self.0.call_mut("Sort", &[columns.into(), sort_object.into()]).unwrap()
    }

    /// Delete a row from the values table by row index or row instance.
    ///
    /// *ОбъектТЗ.Удалить(<Строка>);* or *ОбъектТЗ.Удалить(<Индекс>);*
    ///
    /// Available on a server, thick client, external connection, and a mobile app (server).
    pub fn delete(&mut self, row: ValueTableRowBy) -> SmartVariant {
        self.0.call_mut("Delete", &[row.into()]).unwrap()
    }
}

impl TryFrom<SmartVariant> for ValueTable {
    type Error = ();

    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IDispatch(x) => {
                let result = x.get_ids_of_names(&["ChooseRow", "FindRows"], LOCALE_USER_DEFAULT);
                if SUCCEEDED(result.1) {
                    Ok(ValueTable(x))
                } else {
                    Err(())
                }
            }
            _ => Err(()),
        }
    }
}

impl From<ValueTable> for SmartVariant {
    #[inline]
    fn from(x: ValueTable) -> Self {
        x.0.into()
    }
}

impl ValueTableRow {
    pub fn owner(&self) -> ValueTable {
        self.0.call("Owner", &[]).unwrap().try_into().unwrap()
    }

    pub fn get(&self, index: u32) -> SmartVariant {
        self.0.call("Get", &[SmartVariant::from(index)]).unwrap()
    }

    pub fn set(&mut self, index: u32, value: SmartVariant) -> SmartVariant {
        self.0.call_mut("Set", &[SmartVariant::from(index), value]).unwrap()
    }

    pub fn get_(&self, name: &str) -> SmartVariant {
        self.0.get(&name).unwrap()
    }

    pub fn put(&mut self, name: &str, value: SmartVariant) -> SmartVariant {
        self.0.put(&name, value).unwrap()
    }
}

impl TryFrom<SmartVariant> for ValueTableRow {
    type Error = ();

    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IDispatch(x) => {
                let result = x.get_ids_of_names(&["Owner", "Get", "Set"], LOCALE_USER_DEFAULT);
                if SUCCEEDED(result.1) {
                    Ok(ValueTableRow(x))
                } else {
                    Err(())
                }
            }
            _ => Err(()),
        }
    }
}

impl From<ValueTableRow> for SmartVariant {
    #[inline]
    fn from(x: ValueTableRow) -> Self {
        x.0.into()
    }
}

impl From<ValueTableRowBy> for SmartVariant {
    #[inline]
    fn from(x: ValueTableRowBy) -> Self {
        match x {
            ValueTableRowBy::Index(index) => index.into(),
            ValueTableRowBy::Row(row) => row.into(),
        }
    }
}


impl ValueTableColumn {
    pub fn title(&self) -> String {
        self.0.get("Title").unwrap().try_into().expect("Expected VARIANT containing String!")
    }

    pub fn put_title(&mut self, value: String) -> SmartVariant {
        self.0.put("Title", value.into()).unwrap()
    }

    pub fn name(&self) -> String {
        self.0.get("Name").unwrap().try_into().expect("Expected VARIANT containing String!")
    }

    pub fn put_name(&mut self, value: String) -> SmartVariant {
        self.0.put("Name", value.into()).unwrap()
    }

    pub fn value_type(&self) -> SmartVariant {
        self.0.get("ValueType").unwrap()
    }

    pub fn width(&self) -> u32 {
        self.0.get("Width").unwrap().try_into().expect("Expected VARIANT containing u32 number!")
    }

    pub fn put_width(&mut self, value: u32) -> SmartVariant {
        self.0.put("Name", value.into()).unwrap()
    }
}

impl TryFrom<SmartVariant> for ValueTableColumn {
    type Error = ();

    #[inline]
    fn try_from(x: SmartVariant) -> Result<Self, Self::Error> {
        match x {
            SmartVariant::IDispatch(x) => {
                let result = x.get_ids_of_names(&["Title", "Width", "Name", "ValueType"], LOCALE_USER_DEFAULT);
                if SUCCEEDED(result.1) {
                    Ok(ValueTableColumn(x))
                } else {
                    Err(())
                }
            }
            _ => Err(()),
        }
    }
}

impl From<ValueTableColumn> for SmartVariant {
    #[inline]
    fn from(x: ValueTableColumn) -> Self {
        x.0.into()
    }
}

impl From<ValueTableColumnBy> for SmartVariant {
    #[inline]
    fn from(x: ValueTableColumnBy) -> Self {
        match x {
            ValueTableColumnBy::Index(index) => index.into(),
            ValueTableColumnBy::Name(name) => name.into(),
            ValueTableColumnBy::Column(column) => column.into(),
        }
    }
}

