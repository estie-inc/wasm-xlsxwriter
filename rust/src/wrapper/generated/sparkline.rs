use crate::wrapper::ChartEmptyCells;
use crate::wrapper::ChartRange;
use crate::wrapper::Color;
use crate::wrapper::SparklineType;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Sparkline` struct is used to create an object to represent a sparkline
/// that can be inserted into a worksheet.
///
/// Sparklines are a feature of Excel 2010+ which allows you to add small charts
/// to worksheet cells. These are useful for showing data trends in a compact
/// visual format.
///
/// The following example was used to generate the above file.
///
/// In Excel sparklines can be added as a single entity in a cell that refers to
/// a 1D data range or as a "group" sparkline that is applied across a 1D range
/// and refers to data in a 2D range. A grouped sparkline uses one sparkline for
/// the specified range and any changes to it are applied to the entire
/// sparkline group.
///
/// The {@link Worksheet#addSparkline} method
/// shown allows you to add a sparkline to a single cell that displays data from
/// a 1D range of cells whereas the
/// {@link Worksheet#addSparklineGroup})
/// method applies the group sparkline to a range.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Sparkline {
    pub(crate) inner: Arc<Mutex<xlsx::Sparkline>>,
}

#[wasm_bindgen]
impl Sparkline {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Sparkline {
        Sparkline {
            inner: Arc::new(Mutex::new(xlsx::Sparkline::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> Sparkline {
        Sparkline {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the range of the sparkline data.
    ///
    /// This method is used to set the location of the data from which the
    /// sparkline will be plotted. This constitutes the Y values of the
    /// sparkline.
    ///
    /// By default the X values of the sparkline are taken as evenly spaced
    /// increments. However, it is also possible to specify date values for the
    /// X axis using the
    /// {@link Sparkline#setDateRange}(Sparkline::set_date_range) method.
    ///
    /// The range can either be a 1D range when used with
    /// {@link Worksheet#addSparkline}) or a
    /// 2D range with used with
    /// {@link Worksheet#addSparklineGroup}).
    ///
    /// # Parameters
    ///
    /// - `range`: A 1D or 2D range that contains the data that will be plotted
    ///   in the sparkline. This can specified in different ways, see
    ///   {@link IntoChartRange} for details.
    #[wasm_bindgen(js_name = "setRange", skip_jsdoc)]
    pub fn set_range(&self, range: &ChartRange) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_range(&*range.inner.lock().unwrap());
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the type of sparkline.
    ///
    /// The type of the sparkline can be one of the following:
    ///
    /// - {@link SparklineType#Line}: A line style sparkline. This is the default.
    ///
    /// - {@link SparklineType#Column}: A histogram style sparkline.
    ///
    /// - {@link SparklineType#WinLose}: A positive/negative style sparkline. It
    ///   looks similar to a histogram but all the bars are the same height,
    ///
    /// # Parameters
    ///
    /// - `sparkline_type`: A {@link SparklineType} value.
    #[wasm_bindgen(js_name = "setType", skip_jsdoc)]
    pub fn set_type(&self, sparkline_type: SparklineType) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_type(xlsx::SparklineType::from(sparkline_type));
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the highest point(s) in a sparkline with a marker.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showHighPoint", skip_jsdoc)]
    pub fn show_high_point(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.show_high_point(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the lowest point(s) in a sparkline with a marker.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showLowPoint", skip_jsdoc)]
    pub fn show_low_point(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.show_low_point(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the first point in a sparkline with a marker.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showFirstPoint", skip_jsdoc)]
    pub fn show_first_point(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.show_first_point(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the last point in a sparkline with a marker.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showLastPoint", skip_jsdoc)]
    pub fn show_last_point(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.show_last_point(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the negative points in a sparkline with markers.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showNegativePoints", skip_jsdoc)]
    pub fn show_negative_points(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.show_negative_points(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display markers for all points in the sparkline.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showMarkers", skip_jsdoc)]
    pub fn show_markers(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.show_markers(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the horizontal axis for a sparkline.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showAxis", skip_jsdoc)]
    pub fn show_axis(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.show_axis(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display data from hidden rows or columns in a sparkline.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showHiddenData", skip_jsdoc)]
    pub fn show_hidden_data(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.show_hidden_data(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the option for displaying empty cells in a sparkline.
    ///
    /// The options are:
    ///
    /// - {@link ChartEmptyCells#Gaps}: Show empty cells in the chart as gaps. The
    ///   default.
    /// - {@link ChartEmptyCells#Zero}: Show empty cells in the chart as zeroes.
    /// - {@link ChartEmptyCells#Connected}: Show empty cells in the chart
    ///   connected by a line to the previous point.
    ///
    /// # Parameters
    ///
    /// `option` - A {@link ChartEmptyCells} enum value.
    #[wasm_bindgen(js_name = "showEmptyCellsAs", skip_jsdoc)]
    pub fn show_empty_cells_as(&self, option: ChartEmptyCells) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.show_empty_cells_as(xlsx::ChartEmptyCells::from(option));
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the sparkline in right to left, reversed order.
    ///
    /// Change the default direction of the sparkline so that it is plotted from
    /// right to left instead of the default left to right.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setRightToLeft", skip_jsdoc)]
    pub fn set_right_to_left(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_right_to_left(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the color of a sparkline.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color} such as a html string.
    #[wasm_bindgen(js_name = "setSparklineColor", skip_jsdoc)]
    pub fn set_sparkline_color(&self, color: Color) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_sparkline_color(xlsx::Color::from(color));
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on and set the color the sparkline highest point marker.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color} such as a html string.
    #[wasm_bindgen(js_name = "setHighPointColor", skip_jsdoc)]
    pub fn set_high_point_color(&self, color: Color) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_high_point_color(xlsx::Color::from(color));
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on and set the color the sparkline lowest point marker.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color} such as a html string.
    #[wasm_bindgen(js_name = "setLowPointColor", skip_jsdoc)]
    pub fn set_low_point_color(&self, color: Color) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_low_point_color(xlsx::Color::from(color));
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on and set the color the sparkline first point marker.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color} such as a html string.
    #[wasm_bindgen(js_name = "setFirstPointColor", skip_jsdoc)]
    pub fn set_first_point_color(&self, color: Color) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_first_point_color(xlsx::Color::from(color));
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on and set the color the sparkline last point marker.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color} such as a html string.
    #[wasm_bindgen(js_name = "setLastPointColor", skip_jsdoc)]
    pub fn set_last_point_color(&self, color: Color) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_last_point_color(xlsx::Color::from(color));
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on and set the color the sparkline negative point markers.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color} such as a html string.
    #[wasm_bindgen(js_name = "setNegativePointsColor", skip_jsdoc)]
    pub fn set_negative_points_color(&self, color: Color) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_negative_points_color(xlsx::Color::from(color));
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Turn on and set the color the sparkline point markers.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a {@link Color} enum value or a
    ///   type that can convert `Into` a {@link Color} such as a html string.
    #[wasm_bindgen(js_name = "setMarkersColor", skip_jsdoc)]
    pub fn set_markers_color(&self, color: Color) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_markers_color(xlsx::Color::from(color));
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the weight/width of the sparkline line.
    ///
    /// # Parameters
    ///
    /// - `weight`: The weight/width of the sparkline line. The width can be an
    #[wasm_bindgen(js_name = "setLineWeight", skip_jsdoc)]
    pub fn set_line_weight(&self, weight: f64) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_line_weight(weight);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the maximum vertical value for a sparkline.
    ///
    /// Set the maximum bound to be displayed for the vertical axis of a
    /// sparkline.
    ///
    /// # Parameters
    ///
    /// `max` - The maximum bound for the axes.
    #[wasm_bindgen(js_name = "setCustomMax", skip_jsdoc)]
    pub fn set_custom_max(&self, max: f64) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_custom_max(max);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the minimum vertical value for a sparkline.
    ///
    /// Set the minimum bound to be displayed for the vertical axis of a
    /// sparkline.
    ///
    /// # Parameters
    ///
    /// `min` - The minimum bound for the axes.
    #[wasm_bindgen(js_name = "setCustomMin", skip_jsdoc)]
    pub fn set_custom_min(&self, min: f64) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_custom_min(min);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the maximum vertical value for a group of sparklines.
    ///
    /// Set the maximum vertical value for a group of sparklines based on the
    /// maximum value of the group.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setGroupMax", skip_jsdoc)]
    pub fn set_group_max(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_group_max(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the minimum vertical value for a group of sparklines.
    ///
    /// Set the minimum vertical value for a group of sparklines based on the
    /// minimum value of the group.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setGroupMin", skip_jsdoc)]
    pub fn set_group_min(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_group_min(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the an optional date axis for the sparkline data.
    ///
    /// In general Excel graphs sparklines at equally spaced X intervals.
    /// However, it is also possible to specify an optional range of dates that
    /// can be used as the X values `set_date_range()`.
    ///
    /// # Parameters
    ///
    /// - `range`: A 1D range that contains the dates used to plot the
    ///   sparkline. This can be specified in different ways; see
    ///   {@link IntoChartRange} for details.
    #[wasm_bindgen(js_name = "setDateRange", skip_jsdoc)]
    pub fn set_date_range(&self, range: &ChartRange) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_date_range(&*range.inner.lock().unwrap());
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Change the data range order for 2D data ranges in grouped sparklines.
    ///
    /// When creating grouped sparklines via the
    /// {@link Worksheet#addSparklineGroup})
    /// method, the data range that the sparkline is applied to is in row-major
    /// order, i.e., row by row. If required, you can change this to column-major
    /// order using this method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setColumnOrder", skip_jsdoc)]
    pub fn set_column_order(&self, enable: bool) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_column_order(enable);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the sparkline style type.
    ///
    /// The `set_style()` method is used to set the style of the sparkline to
    /// one of 36 built-in styles. The default style is 1. The image below shows
    /// the 36 default styles. The index is counted from the top left and then
    /// in column-row order.
    ///
    /// # Parameters
    ///
    /// - `style`: An integer value in the range 1-36.
    #[wasm_bindgen(js_name = "setStyle", skip_jsdoc)]
    pub fn set_style(&self, style: u8) -> Sparkline {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::replace(&mut *lock, xlsx::Sparkline::new());
        inner = inner.set_style(style);
        *lock = inner;
        Sparkline {
            inner: Arc::clone(&self.inner),
        }
    }
}
