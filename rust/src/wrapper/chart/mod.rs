// Companion file for Chart: convenience constructors and proxy accessor methods.
// The struct definition and auto-generated methods are in generated/chart.rs.

mod chart_range;

use std::sync::{Arc, Mutex};

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::{Chart, ChartAxis, ChartAxisAccessor, ChartLegend, ChartTitle};

#[wasm_bindgen]
impl Chart {
    #[wasm_bindgen(js_name = "newArea")]
    pub fn new_area() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_area())),
        }
    }

    #[wasm_bindgen(js_name = "newBar")]
    pub fn new_bar() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_bar())),
        }
    }

    #[wasm_bindgen(js_name = "newColumn")]
    pub fn new_column() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_column())),
        }
    }

    #[wasm_bindgen(js_name = "newDoughnut")]
    pub fn new_doughnut() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_doughnut())),
        }
    }

    #[wasm_bindgen(js_name = "newLine")]
    pub fn new_line() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_line())),
        }
    }

    #[wasm_bindgen(js_name = "newPie")]
    pub fn new_pie() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_pie())),
        }
    }

    #[wasm_bindgen(js_name = "newRadar")]
    pub fn new_radar() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_radar())),
        }
    }

    #[wasm_bindgen(js_name = "newScatter")]
    pub fn new_scatter() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_scatter())),
        }
    }

    #[wasm_bindgen(js_name = "newStock")]
    pub fn new_stock() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_stock())),
        }
    }

    #[wasm_bindgen(js_name = "title")]
    pub fn title(&self) -> ChartTitle {
        ChartTitle {
            parent: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "xAxis")]
    pub fn x_axis(&self) -> ChartAxis {
        ChartAxis {
            parent: Arc::clone(&self.inner),
            accessor: ChartAxisAccessor::XAxis,
        }
    }

    #[wasm_bindgen(js_name = "yAxis")]
    pub fn y_axis(&self) -> ChartAxis {
        ChartAxis {
            parent: Arc::clone(&self.inner),
            accessor: ChartAxisAccessor::YAxis,
        }
    }

    #[wasm_bindgen(js_name = "x2Axis")]
    pub fn x2_axis(&self) -> ChartAxis {
        ChartAxis {
            parent: Arc::clone(&self.inner),
            accessor: ChartAxisAccessor::X2Axis,
        }
    }

    #[wasm_bindgen(js_name = "y2Axis")]
    pub fn y2_axis(&self) -> ChartAxis {
        ChartAxis {
            parent: Arc::clone(&self.inner),
            accessor: ChartAxisAccessor::Y2Axis,
        }
    }

    #[wasm_bindgen(js_name = "legend")]
    pub fn legend(&self) -> ChartLegend {
        ChartLegend {
            parent: Arc::clone(&self.inner),
        }
    }
}
