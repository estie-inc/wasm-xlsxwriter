use crate::wrapper::DocProperties;
use crate::wrapper::Format;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Workbook {
    pub(crate) inner: Arc<Mutex<xlsx::Workbook>>,
}

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Workbook {
        Workbook {
            inner: Arc::new(Mutex::new(xlsx::Workbook::new())),
        }
    }
    #[wasm_bindgen(js_name = "addChartsheet", skip_jsdoc)]
    pub fn add_chartsheet(&self) -> &mut Worksheet {
        let lock = self.inner.lock().unwrap();
        lock.add_chartsheet()
    }
    #[wasm_bindgen(js_name = "worksheetFromIndex", skip_jsdoc)]
    pub fn worksheet_from_index(&self, index: usize) -> Result<&mut Worksheet> {
        let lock = self.inner.lock().unwrap();
        lock.worksheet_from_index(index)
    }
    #[wasm_bindgen(js_name = "worksheetFromName", skip_jsdoc)]
    pub fn worksheet_from_name(&self, sheetname: &str) -> Result<&mut Worksheet> {
        let lock = self.inner.lock().unwrap();
        lock.worksheet_from_name(sheetname)
    }
    #[wasm_bindgen(js_name = "defineName", skip_jsdoc)]
    pub fn define_name(&self, name: &str, formula: &str) -> WasmResult<Workbook> {
        let mut lock = self.inner.lock().unwrap();
        lock.define_name(name, formula)?;
        Ok(Workbook {
            inner: Arc::clone(&self.inner),
        })
    }
    #[wasm_bindgen(js_name = "setProperties", skip_jsdoc)]
    pub fn set_properties(&self, properties: DocProperties) -> Workbook {
        let mut lock = self.inner.lock().unwrap();
        lock.set_properties(&properties.inner);
        Workbook {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setDefaultFormat", skip_jsdoc)]
    pub fn set_default_format(
        &self,
        format: Format,
        row_height: u32,
        col_width: u32,
    ) -> WasmResult<Workbook> {
        let mut lock = self.inner.lock().unwrap();
        lock.set_default_format(&format.inner, row_height, col_width)?;
        Ok(Workbook {
            inner: Arc::clone(&self.inner),
        })
    }
    #[wasm_bindgen(js_name = "useExcel2023Theme", skip_jsdoc)]
    pub fn use_excel_2023_theme(&self) -> WasmResult<Workbook> {
        let mut lock = self.inner.lock().unwrap();
        lock.use_excel_2023_theme()?;
        Ok(Workbook {
            inner: Arc::clone(&self.inner),
        })
    }
    #[wasm_bindgen(js_name = "useZipLargeFile", skip_jsdoc)]
    pub fn use_zip_large_file(&self, enable: bool) -> Workbook {
        let mut lock = self.inner.lock().unwrap();
        lock.use_zip_large_file(enable);
        Workbook {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setVbaName", skip_jsdoc)]
    pub fn set_vba_name(&self, name: &str) -> WasmResult<Workbook> {
        let mut lock = self.inner.lock().unwrap();
        lock.set_vba_name(name)?;
        Ok(Workbook {
            inner: Arc::clone(&self.inner),
        })
    }
    #[wasm_bindgen(js_name = "readOnlyRecommended", skip_jsdoc)]
    pub fn read_only_recommended(&self) -> Workbook {
        let mut lock = self.inner.lock().unwrap();
        lock.read_only_recommended();
        Workbook {
            inner: Arc::clone(&self.inner),
        }
    }
}
