use crate::wrapper::ChartGradientFillType;
use crate::wrapper::ChartGradientStop;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartGradientFill` struct represents a gradient fill for a chart
/// element.
///
/// The {@link ChartGradientFill} struct represents the formatting properties for
/// the gradient fill of a Chart element. In Excel a gradient fill is comprised
/// of two or more colors that are blended gradually along a gradient.
///
/// src="https://rustxlsxwriter.github.io/images/gradient_fill_options.png">
///
/// `ChartGradientFill` is a sub property of the {@link ChartFormat} struct and is
/// used with the {@link ChartFormat#setGradientFill} method.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartGradientFill {
    pub(crate) inner: Arc<Mutex<xlsx::ChartGradientFill>>,
}

#[wasm_bindgen]
impl ChartGradientFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartGradientFill {
        ChartGradientFill {
            inner: Arc::new(Mutex::new(xlsx::ChartGradientFill::new())),
        }
    }
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> ChartGradientFill {
        ChartGradientFill {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
    /// Set the type of the gradient fill.
    ///
    /// Change the default type of the gradient fill to one of the styles
    /// supported by Excel.
    ///
    /// The four gradient types supported by Excel are:
    ///
    /// # Parameters
    ///
    /// `gradient_type`: a {@link ChartGradientFillType} enum value.
    #[wasm_bindgen(js_name = "setType", skip_jsdoc)]
    pub fn set_type(&self, gradient_type: ChartGradientFillType) -> ChartGradientFill {
        let mut lock = self.inner.lock().unwrap();
        lock.set_type(xlsx::ChartGradientFillType::from(gradient_type));
        ChartGradientFill {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the gradient stops (data points) for a chart gradient fill.
    ///
    /// A gradient stop, encapsulated by the {@link ChartGradientStop} struct,
    /// represents the properties of a data point that is used to generate a
    /// gradient fill.
    ///
    /// src="https://rustxlsxwriter.github.io/images/gradient_fill_options.png">
    ///
    /// Excel supports between 2 and 10 gradient stops which define the a color
    /// and its position in the gradient as a percentage. These colors and
    /// positions are used to interpolate a gradient fill.
    ///
    /// # Parameters
    ///
    /// `gradient_stops`: A slice ref of {@link ChartGradientStop} values. As in
    /// Excel there must be between 2 and 10 valid gradient stops.
    #[wasm_bindgen(js_name = "setGradientStops", skip_jsdoc)]
    pub fn set_gradient_stops(&self, gradient_stops: Vec<ChartGradientStop>) -> ChartGradientFill {
        let mut lock = self.inner.lock().unwrap();
        lock.set_gradient_stops(
            &gradient_stops
                .iter()
                .map(|x| x.inner.lock().unwrap().clone())
                .collect::<Vec<_>>(),
        );
        ChartGradientFill {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the angle of the linear gradient fill type.
    ///
    /// # Parameters
    ///
    /// - `angle`: The angle of the linear gradient fill in the range `0 <=
    ///   angle < 360`. The default angle is 90 degrees.
    #[wasm_bindgen(js_name = "setAngle", skip_jsdoc)]
    pub fn set_angle(&self, angle: u16) -> ChartGradientFill {
        let mut lock = self.inner.lock().unwrap();
        lock.set_angle(angle);
        ChartGradientFill {
            inner: Arc::clone(&self.inner),
        }
    }
}
