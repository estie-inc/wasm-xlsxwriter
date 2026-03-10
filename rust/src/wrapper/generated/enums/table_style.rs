use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `TableStyle` enum defines the worksheet table styles.
///
/// Excel supports 61 different styles for tables divided into Light, Medium and
/// Dark categories. You can set one of these styles using a `TableStyle` enum
/// value.
///
/// The style is set via the {@link Table#setStyle} method. The default table
/// style in Excel is equivalent to {@link TableStyle#Medium9}.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum TableStyle {
    /// No table style.
    None,
    /// Table Style Light 1, White.
    Light1,
    /// Table Style Light 2, Light Blue.
    Light2,
    /// Table Style Light 3, Light Orange.
    Light3,
    /// Table Style Light 4, White.
    Light4,
    /// Table Style Light 5, Light Yellow.
    Light5,
    /// Table Style Light 6, Light Blue.
    Light6,
    /// Table Style Light 7, Light Green.
    Light7,
    /// Table Style Light 8, White.
    Light8,
    /// Table Style Light 9, Blue.
    Light9,
    /// Table Style Light 10, Orange.
    Light10,
    /// Table Style Light 11, White.
    Light11,
    /// Table Style Light 12, Gold.
    Light12,
    /// Table Style Light 13, Blue.
    Light13,
    /// Table Style Light 14, Green.
    Light14,
    /// Table Style Light 15, White.
    Light15,
    /// Table Style Light 16, Light Blue.
    Light16,
    /// Table Style Light 17, Light Orange.
    Light17,
    /// Table Style Light 18, White.
    Light18,
    /// Table Style Light 19, Light Yellow.
    Light19,
    /// Table Style Light 20, Light Blue.
    Light20,
    /// Table Style Light 21, Light Green.
    Light21,
    /// Table Style Medium 1, White.
    Medium1,
    /// Table Style Medium 2, Blue.
    Medium2,
    /// Table Style Medium 3, Orange.
    Medium3,
    /// Table Style Medium 4, White.
    Medium4,
    /// Table Style Medium 5, Gold.
    Medium5,
    /// Table Style Medium 6, Blue.
    Medium6,
    /// Table Style Medium 7, Green.
    Medium7,
    /// Table Style Medium 8, Light Grey.
    Medium8,
    /// Table Style Medium 9, Blue.
    Medium9,
    /// Table Style Medium 10, Orange.
    Medium10,
    /// Table Style Medium 11, Light Grey.
    Medium11,
    /// Table Style Medium 12, Gold.
    Medium12,
    /// Table Style Medium 13, Blue.
    Medium13,
    /// Table Style Medium 14, Green.
    Medium14,
    /// Table Style Medium 15, White.
    Medium15,
    /// Table Style Medium 16, Blue.
    Medium16,
    /// Table Style Medium 17, Orange.
    Medium17,
    /// Table Style Medium 18, White.
    Medium18,
    /// Table Style Medium 19, Gold.
    Medium19,
    /// Table Style Medium 20, Blue.
    Medium20,
    /// Table Style Medium 21, Green.
    Medium21,
    /// Table Style Medium 22, Light Grey.
    Medium22,
    /// Table Style Medium 23, Light Blue.
    Medium23,
    /// Table Style Medium 24, Light Orange.
    Medium24,
    /// Table Style Medium 25, Light Grey.
    Medium25,
    /// Table Style Medium 26, Light Yellow.
    Medium26,
    /// Table Style Medium 27, Light Blue.
    Medium27,
    /// Table Style Medium 28, Light Green.
    Medium28,
    /// Table Style Dark 1, Dark Grey.
    Dark1,
    /// Table Style Dark 2, Dark Blue.
    Dark2,
    /// Table Style Dark 3, Brown.
    Dark3,
    /// Table Style Dark 4, Grey.
    Dark4,
    /// Table Style Dark 5, Dark Yellow.
    Dark5,
    /// Table Style Dark 6, Blue.
    Dark6,
    /// Table Style Dark 7, Dark Green.
    Dark7,
    /// Table Style Dark 8, Light Grey.
    Dark8,
    /// Table Style Dark 9, Light Orange.
    Dark9,
    /// Table Style Dark 10, Gold.
    Dark10,
    /// Table Style Dark 11, Green.
    Dark11,
}

impl From<TableStyle> for xlsx::TableStyle {
    fn from(value: TableStyle) -> xlsx::TableStyle {
        match value {
            TableStyle::None => xlsx::TableStyle::None,
            TableStyle::Light1 => xlsx::TableStyle::Light1,
            TableStyle::Light2 => xlsx::TableStyle::Light2,
            TableStyle::Light3 => xlsx::TableStyle::Light3,
            TableStyle::Light4 => xlsx::TableStyle::Light4,
            TableStyle::Light5 => xlsx::TableStyle::Light5,
            TableStyle::Light6 => xlsx::TableStyle::Light6,
            TableStyle::Light7 => xlsx::TableStyle::Light7,
            TableStyle::Light8 => xlsx::TableStyle::Light8,
            TableStyle::Light9 => xlsx::TableStyle::Light9,
            TableStyle::Light10 => xlsx::TableStyle::Light10,
            TableStyle::Light11 => xlsx::TableStyle::Light11,
            TableStyle::Light12 => xlsx::TableStyle::Light12,
            TableStyle::Light13 => xlsx::TableStyle::Light13,
            TableStyle::Light14 => xlsx::TableStyle::Light14,
            TableStyle::Light15 => xlsx::TableStyle::Light15,
            TableStyle::Light16 => xlsx::TableStyle::Light16,
            TableStyle::Light17 => xlsx::TableStyle::Light17,
            TableStyle::Light18 => xlsx::TableStyle::Light18,
            TableStyle::Light19 => xlsx::TableStyle::Light19,
            TableStyle::Light20 => xlsx::TableStyle::Light20,
            TableStyle::Light21 => xlsx::TableStyle::Light21,
            TableStyle::Medium1 => xlsx::TableStyle::Medium1,
            TableStyle::Medium2 => xlsx::TableStyle::Medium2,
            TableStyle::Medium3 => xlsx::TableStyle::Medium3,
            TableStyle::Medium4 => xlsx::TableStyle::Medium4,
            TableStyle::Medium5 => xlsx::TableStyle::Medium5,
            TableStyle::Medium6 => xlsx::TableStyle::Medium6,
            TableStyle::Medium7 => xlsx::TableStyle::Medium7,
            TableStyle::Medium8 => xlsx::TableStyle::Medium8,
            TableStyle::Medium9 => xlsx::TableStyle::Medium9,
            TableStyle::Medium10 => xlsx::TableStyle::Medium10,
            TableStyle::Medium11 => xlsx::TableStyle::Medium11,
            TableStyle::Medium12 => xlsx::TableStyle::Medium12,
            TableStyle::Medium13 => xlsx::TableStyle::Medium13,
            TableStyle::Medium14 => xlsx::TableStyle::Medium14,
            TableStyle::Medium15 => xlsx::TableStyle::Medium15,
            TableStyle::Medium16 => xlsx::TableStyle::Medium16,
            TableStyle::Medium17 => xlsx::TableStyle::Medium17,
            TableStyle::Medium18 => xlsx::TableStyle::Medium18,
            TableStyle::Medium19 => xlsx::TableStyle::Medium19,
            TableStyle::Medium20 => xlsx::TableStyle::Medium20,
            TableStyle::Medium21 => xlsx::TableStyle::Medium21,
            TableStyle::Medium22 => xlsx::TableStyle::Medium22,
            TableStyle::Medium23 => xlsx::TableStyle::Medium23,
            TableStyle::Medium24 => xlsx::TableStyle::Medium24,
            TableStyle::Medium25 => xlsx::TableStyle::Medium25,
            TableStyle::Medium26 => xlsx::TableStyle::Medium26,
            TableStyle::Medium27 => xlsx::TableStyle::Medium27,
            TableStyle::Medium28 => xlsx::TableStyle::Medium28,
            TableStyle::Dark1 => xlsx::TableStyle::Dark1,
            TableStyle::Dark2 => xlsx::TableStyle::Dark2,
            TableStyle::Dark3 => xlsx::TableStyle::Dark3,
            TableStyle::Dark4 => xlsx::TableStyle::Dark4,
            TableStyle::Dark5 => xlsx::TableStyle::Dark5,
            TableStyle::Dark6 => xlsx::TableStyle::Dark6,
            TableStyle::Dark7 => xlsx::TableStyle::Dark7,
            TableStyle::Dark8 => xlsx::TableStyle::Dark8,
            TableStyle::Dark9 => xlsx::TableStyle::Dark9,
            TableStyle::Dark10 => xlsx::TableStyle::Dark10,
            TableStyle::Dark11 => xlsx::TableStyle::Dark11,
        }
    }
}
