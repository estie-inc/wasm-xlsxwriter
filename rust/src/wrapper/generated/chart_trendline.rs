use crate::wrapper::ChartFont;
use crate::wrapper::ChartFormat;
use crate::wrapper::ChartTrendlineType;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartTrendline` struct represents a trendline for a chart series.
///
/// Excel allows you to add a trendline to a data series that represents the
/// trend or regression of the data using different types of fit. The
/// `ChartTrendline` struct represents the options of Excel trendlines and can
/// be added to a series via the {@link ChartSeries#setTrendline} method.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartTrendline {
    pub(crate) inner: Arc<Mutex<xlsx::ChartTrendline>>,
}

#[wasm_bindgen]
impl ChartTrendline {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartTrendline {
        ChartTrendline {
            inner: Arc::new(Mutex::new(xlsx::ChartTrendline::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ChartTrendline {
        ChartTrendline {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the type of the Chart series trendlines.
    ///
    /// Set the trendline type to one of the Excel allowable types represented
    /// by the {@link ChartTrendlineType} enum.
    ///
    /// # Parameters
    ///
    /// - `trend`: A {@link ChartTrendlineType} enum reference.
    #[wasm_bindgen(js_name = "setType", skip_jsdoc)]
    pub fn set_type(&self, trend: ChartTrendlineType) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_type(xlsx::ChartTrendlineType::from(trend));
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the formatting properties for a chart trendline.
    ///
    /// Set the formatting properties for a chart trendline via a
    /// {@link ChartFormat} object or a sub struct that implements
    /// {@link IntoChartFormat}.
    ///
    /// The formatting that can be applied via a {@link ChartFormat} object are:
    /// - {@link ChartFormat#setSolidFill}: Set the {@link ChartSolidFill} properties.
    /// - {@link ChartFormat#setPatternFill}: Set the {@link ChartPatternFill} properties.
    /// - {@link ChartFormat#setGradientFill}: Set the {@link ChartGradientFill} properties.
    /// - {@link ChartFormat#setNoFill}: Turn off the fill for the chart object.
    /// - {@link ChartFormat#setLine}: Set the {@link ChartLine} properties.
    /// - {@link ChartFormat#setBorder}: Set the {@link ChartBorder} properties.
    ///   A synonym for {@link ChartLine} depending on context.
    /// - {@link ChartFormat#setNoLine}: Turn off the line for the chart object.
    /// - {@link ChartFormat#setNoBorder}: Turn off the border for the chart object.
    /// - `set_no_border`: Turn off the border for the chart object.
    ///
    /// # Parameters
    ///
    /// `format`: A {@link ChartFormat} struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// {@link IntoChartFormat} for details.
    #[wasm_bindgen(js_name = "setFormat", skip_jsdoc)]
    pub fn set_format(&self, format: &ChartFormat) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_format(&mut *format.inner.lock().unwrap());
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the formatting properties for a chart trendline label.
    ///
    /// Set the formatting properties for a chart trendline label via a
    /// {@link ChartFormat} object or a sub struct that implements
    /// {@link IntoChartFormat}. The label is displayed when you use the
    /// {@link ChartTrendline#displayEquation} or
    /// {@link ChartTrendline#displayRSquared} methods.
    ///
    /// The formatting that can be applied via a {@link ChartFormat} object are:
    ///
    /// - {@link ChartFormat#setSolidFill}: Set the {@link ChartSolidFill} properties.
    /// - {@link ChartFormat#setPatternFill}: Set the {@link ChartPatternFill} properties.
    /// - {@link ChartFormat#setGradientFill}: Set the {@link ChartGradientFill} properties.
    /// - {@link ChartFormat#setNoFill}: Turn off the fill for the chart object.
    /// - {@link ChartFormat#setLine}: Set the {@link ChartLine} properties.
    /// - {@link ChartFormat#setBorder}: Set the {@link ChartBorder} properties.
    ///   A synonym for {@link ChartLine} depending on context.
    /// - {@link ChartFormat#setNoLine}: Turn off the line for the chart object.
    /// - {@link ChartFormat#setNoBorder}: Turn off the border for the chart object.
    ///
    /// # Parameters
    ///
    /// `format`: A {@link ChartFormat} struct reference or a sub struct that will
    /// convert into a `ChartFormat` instance. See the docs for
    /// {@link IntoChartFormat} for details.
    #[wasm_bindgen(js_name = "setLabelFormat", skip_jsdoc)]
    pub fn set_label_format(&self, format: &ChartFormat) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_label_format(&mut *format.inner.lock().unwrap());
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the font properties of a chart trendline label.
    ///
    /// Set the font properties of a chart trendline label using a {@link ChartFont}
    /// reference. The label is displayed when you use the
    /// {@link ChartTrendline#displayEquation} or
    /// {@link ChartTrendline#displayRSquared} methods.
    ///
    /// Example font properties that can be set are:
    ///
    /// - {@link ChartFont#setBold}
    /// - {@link ChartFont#setItalic}
    /// - {@link ChartFont#setName}
    /// - {@link ChartFont#setSize}
    /// - {@link ChartFont#setRotation}
    ///
    /// See {@link ChartFont} for full details.
    ///
    /// # Parameters
    ///
    /// `font`: A {@link ChartFont} struct reference to represent the font
    /// properties.
    #[wasm_bindgen(js_name = "setLabelFont", skip_jsdoc)]
    pub fn set_label_font(&self, font: &ChartFont) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_label_font(&*font.inner.lock().unwrap());
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the name for a chart trendline.
    ///
    /// Set a custom name for a the trendline when it is displayed in the chart
    /// legend.
    ///
    /// # Parameters
    ///
    /// - `name`: The custom string to name the trendline in the chart legend.
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_name(name);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the forward period for a chart trendline.
    ///
    /// Extend the trendline forward by a multiplier of the default length.
    ///
    /// # Parameters
    ///
    /// - `period`: The forward period value.
    #[wasm_bindgen(js_name = "setForwardPeriod", skip_jsdoc)]
    pub fn set_forward_period(&self, period: f64) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_forward_period(period);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the backward period for a chart trendline.
    ///
    /// Extend the trendline backward by a multiplier of the default length.
    ///
    /// # Parameters
    ///
    /// - `period`: The backward period value.
    #[wasm_bindgen(js_name = "setBackwardPeriod", skip_jsdoc)]
    pub fn set_backward_period(&self, period: f64) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_backward_period(period);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the trendline equation for a chart trendline.
    ///
    /// Note, the equation is calculated by Excel at runtime. It isn't
    /// calculated by `rust_xlsxwriter` or stored in the Excel file format.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "displayEquation", skip_jsdoc)]
    pub fn display_equation(&self, enable: bool) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.display_equation(enable);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display the R-squared value for a chart trendline.
    ///
    /// Display the R-squared [coefficient of determination] for the trendline
    /// as an indicator of how accurate the fit is.
    ///
    /// Note, the R-squared value is calculated by Excel at runtime. It isn't
    /// calculated by `rust_xlsxwriter` or stored in the Excel file format.
    ///
    /// [coefficient of determination]:
    ///     https://en.wikipedia.org/wiki/Coefficient_of_determination
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "displayRSquared", skip_jsdoc)]
    pub fn display_r_squared(&self, enable: bool) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.display_r_squared(enable);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Y-axis intercept for a chart trendline.
    ///
    /// Set the point where the trendline will intercept the Y-axis.
    ///
    /// # Parameters
    ///
    /// - `intercept`: The intercept with the Y-axis.
    #[wasm_bindgen(js_name = "setIntercept", skip_jsdoc)]
    pub fn set_intercept(&self, intercept: f64) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.set_intercept(intercept);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Delete/hide the trendline name from the chart legend.
    ///
    /// The `delete_from_legend()` method deletes/hides the trendline name from
    /// the chart legend. This is often desirable since the trendlines are
    /// generally obvious relative to their series and their names can clutter
    /// the chart legend.
    ///
    /// See also the {@link ChartSeries#deleteFromLegend} and the
    /// {@link ChartLegend#deleteEntries} methods.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "deleteFromLegend", skip_jsdoc)]
    pub fn delete_from_legend(&self, enable: bool) -> ChartTrendline {
        let mut lock = self.inner.lock().unwrap();
        lock.delete_from_legend(enable);
        ChartTrendline {
            inner: Arc::clone(&self.inner),
        }
    }
}
