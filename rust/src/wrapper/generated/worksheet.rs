use crate::wrapper::Format;
use crate::wrapper::Image;
use crate::wrapper::ProtectionOptions;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone, Copy)]
pub enum WorksheetAccessor {
    AddWorksheet,
    AddChartsheet,
}

#[derive(Clone)]
#[wasm_bindgen]
pub struct Worksheet {
    pub(crate) parent: Arc<Mutex<xlsx::Workbook>>,
    pub(crate) accessor: WorksheetAccessor,
}

#[wasm_bindgen]
impl Worksheet {
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_name(name),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_name(name),
        }?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
        })
    }
    #[wasm_bindgen(js_name = "name", skip_jsdoc)]
    pub fn name(&self) -> String {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().name(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().name(),
        }
    }
    #[wasm_bindgen(js_name = "insertBackgroundImage", skip_jsdoc)]
    pub fn insert_background_image(&self, image: Image) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().insert_background_image(&image.inner)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().insert_background_image(&image.inner)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "showAllNotes", skip_jsdoc)]
    pub fn show_all_notes(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().show_all_notes(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().show_all_notes(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setDefaultNoteAuthor", skip_jsdoc)]
    pub fn set_default_note_author(&self, name: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_default_note_author(name),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_default_note_author(name),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "groupSymbolsAbove", skip_jsdoc)]
    pub fn group_symbols_above(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().group_symbols_above(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().group_symbols_above(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "groupSymbolsToLeft", skip_jsdoc)]
    pub fn group_symbols_to_left(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().group_symbols_to_left(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().group_symbols_to_left(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setDefaultRowHeight", skip_jsdoc)]
    pub fn set_default_row_height(&self, height: f64) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_default_row_height(height),
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_default_row_height(height)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setDefaultRowHeightPixels", skip_jsdoc)]
    pub fn set_default_row_height_pixels(&self, height: u32) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_default_row_height_pixels(height)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_default_row_height_pixels(height)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "hideUnusedRows", skip_jsdoc)]
    pub fn hide_unused_rows(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().hide_unused_rows(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().hide_unused_rows(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setDefaultFormat", skip_jsdoc)]
    pub fn set_default_format(
        &self,
        format: Format,
        row_height: u32,
        col_width: u32,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet()
                    .set_default_format(&format.inner, row_height, col_width)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet()
                    .set_default_format(&format.inner, row_height, col_width)
            }
        }?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
        })
    }
    #[wasm_bindgen(js_name = "filterAutomaticOff", skip_jsdoc)]
    pub fn filter_automatic_off(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().filter_automatic_off(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().filter_automatic_off(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "protect", skip_jsdoc)]
    pub fn protect(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().protect(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().protect(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "protectWithPassword", skip_jsdoc)]
    pub fn protect_with_password(&self, password: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().protect_with_password(password),
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().protect_with_password(password)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "protectWithOptions", skip_jsdoc)]
    pub fn protect_with_options(&self, options: ProtectionOptions) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .protect_with_options(&*options.inner.lock().unwrap()),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .protect_with_options(&*options.inner.lock().unwrap()),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setFormulaResultDefault", skip_jsdoc)]
    pub fn set_formula_result_default(&self, result: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_formula_result_default(result)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_formula_result_default(result)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setRightToLeft", skip_jsdoc)]
    pub fn set_right_to_left(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_right_to_left(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_right_to_left(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setActive", skip_jsdoc)]
    pub fn set_active(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_active(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_active(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setSelected", skip_jsdoc)]
    pub fn set_selected(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_selected(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_selected(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_hidden(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_hidden(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setVeryHidden", skip_jsdoc)]
    pub fn set_very_hidden(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_very_hidden(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_very_hidden(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setFirstTab", skip_jsdoc)]
    pub fn set_first_tab(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_first_tab(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_first_tab(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPaperSize", skip_jsdoc)]
    pub fn set_paper_size(&self, paper_size: u8) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_paper_size(paper_size),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_paper_size(paper_size),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPageOrder", skip_jsdoc)]
    pub fn set_page_order(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_page_order(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_page_order(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setLandscape", skip_jsdoc)]
    pub fn set_landscape(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_landscape(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_landscape(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPortrait", skip_jsdoc)]
    pub fn set_portrait(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_portrait(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_portrait(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setViewNormal", skip_jsdoc)]
    pub fn set_view_normal(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_view_normal(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_view_normal(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setViewPageLayout", skip_jsdoc)]
    pub fn set_view_page_layout(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_view_page_layout(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_view_page_layout(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setViewPageBreakPreview", skip_jsdoc)]
    pub fn set_view_page_break_preview(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_view_page_break_preview(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_view_page_break_preview(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setVerticalPageBreaks", skip_jsdoc)]
    pub fn set_vertical_page_breaks(&self, breaks: Vec<u32>) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_vertical_page_breaks(&breaks)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_vertical_page_breaks(&breaks)
            }
        }?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
        })
    }
    #[wasm_bindgen(js_name = "setZoom", skip_jsdoc)]
    pub fn set_zoom(&self, zoom: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_zoom(zoom),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_zoom(zoom),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setZoomToFit", skip_jsdoc)]
    pub fn set_zoom_to_fit(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_zoom_to_fit(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_zoom_to_fit(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setHeader", skip_jsdoc)]
    pub fn set_header(&self, header: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_header(header),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_header(header),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setFooter", skip_jsdoc)]
    pub fn set_footer(&self, footer: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_footer(footer),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_footer(footer),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setHeaderFooterScaleWithDoc", skip_jsdoc)]
    pub fn set_header_footer_scale_with_doc(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .set_header_footer_scale_with_doc(enable),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .set_header_footer_scale_with_doc(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setHeaderFooterAlignWithPage", skip_jsdoc)]
    pub fn set_header_footer_align_with_page(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .set_header_footer_align_with_page(enable),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .set_header_footer_align_with_page(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMargins", skip_jsdoc)]
    pub fn set_margins(
        &self,
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .set_margins(left, right, top, bottom, header, footer),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .set_margins(left, right, top, bottom, header, footer),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPrintFirstPageNumber", skip_jsdoc)]
    pub fn set_print_first_page_number(&self, page_number: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock
                .add_worksheet()
                .set_print_first_page_number(page_number),
            WorksheetAccessor::AddChartsheet => lock
                .add_chartsheet()
                .set_print_first_page_number(page_number),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPrintScale", skip_jsdoc)]
    pub fn set_print_scale(&self, scale: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_print_scale(scale),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_print_scale(scale),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPrintFitToPages", skip_jsdoc)]
    pub fn set_print_fit_to_pages(&self, width: u16, height: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_print_fit_to_pages(width, height)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_print_fit_to_pages(width, height)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPrintCenterHorizontally", skip_jsdoc)]
    pub fn set_print_center_horizontally(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_print_center_horizontally(enable)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_print_center_horizontally(enable)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPrintCenterVertically", skip_jsdoc)]
    pub fn set_print_center_vertically(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_print_center_vertically(enable)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_print_center_vertically(enable)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setScreenGridlines", skip_jsdoc)]
    pub fn set_screen_gridlines(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_screen_gridlines(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_screen_gridlines(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPrintGridlines", skip_jsdoc)]
    pub fn set_print_gridlines(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_print_gridlines(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_print_gridlines(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPrintBlackAndWhite", skip_jsdoc)]
    pub fn set_print_black_and_white(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_print_black_and_white(enable)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_print_black_and_white(enable)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPrintDraft", skip_jsdoc)]
    pub fn set_print_draft(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_print_draft(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_print_draft(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPrintHeadings", skip_jsdoc)]
    pub fn set_print_headings(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_print_headings(enable),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_print_headings(enable),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "autofit", skip_jsdoc)]
    pub fn autofit(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().autofit(),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().autofit(),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setAutofitMaxWidth", skip_jsdoc)]
    pub fn set_autofit_max_width(&self, max_width: u32) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => {
                lock.add_worksheet().set_autofit_max_width(max_width)
            }
            WorksheetAccessor::AddChartsheet => {
                lock.add_chartsheet().set_autofit_max_width(max_width)
            }
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setVbaName", skip_jsdoc)]
    pub fn set_vba_name(&self, name: &str) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_vba_name(name),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_vba_name(name),
        }?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
        })
    }
    #[wasm_bindgen(js_name = "setNanValue", skip_jsdoc)]
    pub fn set_nan_value(&self, value: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_nan_value(value),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_nan_value(value),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setInfinityValue", skip_jsdoc)]
    pub fn set_infinity_value(&self, value: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_infinity_value(value),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_infinity_value(value),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setNegInfinityValue", skip_jsdoc)]
    pub fn set_neg_infinity_value(&self, value: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            WorksheetAccessor::AddWorksheet => lock.add_worksheet().set_neg_infinity_value(value),
            WorksheetAccessor::AddChartsheet => lock.add_chartsheet().set_neg_infinity_value(value),
        }
        Worksheet {
            parent: Arc::clone(&self.parent),
        }
    }
}
