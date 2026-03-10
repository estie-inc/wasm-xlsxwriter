use crate::wrapper::ChartAxisCrossing;
use crate::wrapper::ChartAxisDateUnitType;
use crate::wrapper::ChartAxisDisplayUnitType;
use crate::wrapper::ChartAxisLabelAlignment;
use crate::wrapper::ChartAxisLabelPosition;
use crate::wrapper::ChartAxisTickType;
use crate::wrapper::ChartFont;
use crate::wrapper::ChartLayout;
use crate::wrapper::ChartLine;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone, Copy)]
pub enum ChartAxisAccessor {
    XAxis,
    YAxis,
    X2Axis,
    Y2Axis,
}

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartAxis {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
    pub(crate) accessor: ChartAxisAccessor,
}

#[wasm_bindgen]
impl ChartAxis {
    #[wasm_bindgen(js_name = "setNameFont", skip_jsdoc)]
    pub fn set_name_font(&self, font: ChartFont) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_name_font(&font.inner),
            ChartAxisAccessor::YAxis => lock.y_axis().set_name_font(&font.inner),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_name_font(&font.inner),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_name_font(&font.inner),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: ChartFont) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_font(&font.inner),
            ChartAxisAccessor::YAxis => lock.y_axis().set_font(&font.inner),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_font(&font.inner),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_font(&font.inner),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setNumFormat", skip_jsdoc)]
    pub fn set_num_format(&self, num_format: &str) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_num_format(num_format),
            ChartAxisAccessor::YAxis => lock.y_axis().set_num_format(num_format),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_num_format(num_format),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_num_format(num_format),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setDateAxis", skip_jsdoc)]
    pub fn set_date_axis(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_date_axis(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_date_axis(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_date_axis(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_date_axis(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setTextAxis", skip_jsdoc)]
    pub fn set_text_axis(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_text_axis(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_text_axis(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_text_axis(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_text_axis(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setAutomaticAxis", skip_jsdoc)]
    pub fn set_automatic_axis(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_automatic_axis(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_automatic_axis(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_automatic_axis(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_automatic_axis(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setCrossing", skip_jsdoc)]
    pub fn set_crossing(&self, crossing: ChartAxisCrossing) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_crossing(xlsx::ChartAxisCrossing::from(crossing)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_crossing(xlsx::ChartAxisCrossing::from(crossing)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_crossing(xlsx::ChartAxisCrossing::from(crossing)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_crossing(xlsx::ChartAxisCrossing::from(crossing)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setReverse", skip_jsdoc)]
    pub fn set_reverse(&self) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_reverse(),
            ChartAxisAccessor::YAxis => lock.y_axis().set_reverse(),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_reverse(),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_reverse(),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMax", skip_jsdoc)]
    pub fn set_max(&self, max: f64) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_max(max),
            ChartAxisAccessor::YAxis => lock.y_axis().set_max(max),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_max(max),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_max(max),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMin", skip_jsdoc)]
    pub fn set_min(&self, min: f64) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_min(min),
            ChartAxisAccessor::YAxis => lock.y_axis().set_min(min),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_min(min),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_min(min),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMajorUnit", skip_jsdoc)]
    pub fn set_major_unit(&self, value: f64) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_major_unit(value),
            ChartAxisAccessor::YAxis => lock.y_axis().set_major_unit(value),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_major_unit(value),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_major_unit(value),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMinorUnit", skip_jsdoc)]
    pub fn set_minor_unit(&self, value: f64) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_minor_unit(value),
            ChartAxisAccessor::YAxis => lock.y_axis().set_minor_unit(value),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_minor_unit(value),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_minor_unit(value),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setDisplayUnitType", skip_jsdoc)]
    pub fn set_display_unit_type(&self, unit_type: ChartAxisDisplayUnitType) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_display_unit_type(xlsx::ChartAxisDisplayUnitType::from(unit_type)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_display_unit_type(xlsx::ChartAxisDisplayUnitType::from(unit_type)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_display_unit_type(xlsx::ChartAxisDisplayUnitType::from(unit_type)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_display_unit_type(xlsx::ChartAxisDisplayUnitType::from(unit_type)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setDisplayUnitsVisible", skip_jsdoc)]
    pub fn set_display_units_visible(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_display_units_visible(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_display_units_visible(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_display_units_visible(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_display_units_visible(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMajorUnitDateType", skip_jsdoc)]
    pub fn set_major_unit_date_type(&self, unit_type: ChartAxisDateUnitType) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_major_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_major_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_major_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_major_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMinorUnitDateType", skip_jsdoc)]
    pub fn set_minor_unit_date_type(&self, unit_type: ChartAxisDateUnitType) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_minor_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_minor_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_minor_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_minor_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setLabelAlignment", skip_jsdoc)]
    pub fn set_label_alignment(&self, alignment: ChartAxisLabelAlignment) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_label_alignment(xlsx::ChartAxisLabelAlignment::from(alignment)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_label_alignment(xlsx::ChartAxisLabelAlignment::from(alignment)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_label_alignment(xlsx::ChartAxisLabelAlignment::from(alignment)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_label_alignment(xlsx::ChartAxisLabelAlignment::from(alignment)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMajorGridlines", skip_jsdoc)]
    pub fn set_major_gridlines(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_major_gridlines(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_major_gridlines(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_major_gridlines(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_major_gridlines(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMinorGridlines", skip_jsdoc)]
    pub fn set_minor_gridlines(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_minor_gridlines(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_minor_gridlines(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_minor_gridlines(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_minor_gridlines(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMajorGridlinesLine", skip_jsdoc)]
    pub fn set_major_gridlines_line(&self, line: ChartLine) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_major_gridlines_line(&line.inner),
            ChartAxisAccessor::YAxis => lock.y_axis().set_major_gridlines_line(&line.inner),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_major_gridlines_line(&line.inner),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_major_gridlines_line(&line.inner),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMinorGridlinesLine", skip_jsdoc)]
    pub fn set_minor_gridlines_line(&self, line: ChartLine) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_minor_gridlines_line(&line.inner),
            ChartAxisAccessor::YAxis => lock.y_axis().set_minor_gridlines_line(&line.inner),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_minor_gridlines_line(&line.inner),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_minor_gridlines_line(&line.inner),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setLabelPosition", skip_jsdoc)]
    pub fn set_label_position(&self, position: ChartAxisLabelPosition) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_label_position(xlsx::ChartAxisLabelPosition::from(position)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_label_position(xlsx::ChartAxisLabelPosition::from(position)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_label_position(xlsx::ChartAxisLabelPosition::from(position)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_label_position(xlsx::ChartAxisLabelPosition::from(position)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPositionBetweenTicks", skip_jsdoc)]
    pub fn set_position_between_ticks(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_position_between_ticks(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_position_between_ticks(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_position_between_ticks(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_position_between_ticks(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setLabelInterval", skip_jsdoc)]
    pub fn set_label_interval(&self, interval: u16) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_label_interval(interval),
            ChartAxisAccessor::YAxis => lock.y_axis().set_label_interval(interval),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_label_interval(interval),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_label_interval(interval),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setTickInterval", skip_jsdoc)]
    pub fn set_tick_interval(&self, interval: u16) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_tick_interval(interval),
            ChartAxisAccessor::YAxis => lock.y_axis().set_tick_interval(interval),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_tick_interval(interval),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_tick_interval(interval),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMajorTickType", skip_jsdoc)]
    pub fn set_major_tick_type(&self, tick_type: ChartAxisTickType) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_major_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_major_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_major_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_major_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMinorTickType", skip_jsdoc)]
    pub fn set_minor_tick_type(&self, tick_type: ChartAxisTickType) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_minor_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_minor_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_minor_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_minor_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setLogBase", skip_jsdoc)]
    pub fn set_log_base(&self, base: u16) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_log_base(base),
            ChartAxisAccessor::YAxis => lock.y_axis().set_log_base(base),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_log_base(base),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_log_base(base),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_hidden(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_hidden(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_hidden(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_hidden(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setLabelLayout", skip_jsdoc)]
    pub fn set_label_layout(&self, layout: ChartLayout) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_label_layout(&layout.inner),
            ChartAxisAccessor::YAxis => lock.y_axis().set_label_layout(&layout.inner),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_label_layout(&layout.inner),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_label_layout(&layout.inner),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
}
