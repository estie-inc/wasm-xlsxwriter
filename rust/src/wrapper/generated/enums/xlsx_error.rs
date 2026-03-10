use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `XlsxError` enum defines the error values for the `rust_xlsxwriter`
/// library.
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum XlsxError {
    /// A general parameter error that is raised when a parameter conflicts with
    /// an Excel limit or syntax. The nature of the error is in the error string.
    ParameterError(String),
    /// Row or column argument exceeds Excel's limits of 1,048,576 rows and
    /// 16,384 columns for a worksheet.
    RowColumnLimitError,
    /// The first row or column is greater than the last row or column in a range
    /// specification, i.e., the order is reversed.
    RowColumnOrderError,
    /// Worksheet name cannot be blank.
    SheetnameCannotBeBlank(String),
    /// Worksheet name exceeds Excel's limit of 31 characters.
    SheetnameLengthExceeded(String),
    /// Worksheet name is already in use in the workbook.
    SheetnameReused(String),
    /// Worksheet name cannot contain any of the following invalid characters: `[ ] : * ? / \`
    SheetnameContainsInvalidCharacter(String),
    /// Worksheet name cannot start or end with an apostrophe.
    SheetnameStartsOrEndsWithApostrophe(String),
    /// String exceeds Excel's limit of 32,767 characters.
    MaxStringLengthExceeded,
    /// Error when trying to retrieve a worksheet reference by index or by name.
    UnknownWorksheetNameOrIndex(String),
    /// A merge range cannot be a single cell in Excel.
    MergeRangeSingleCell,
    /// The merge range overlaps a previous merge range. This is strictly
    /// prohibited by Excel.
    MergeRangeOverlaps(String, String),
    /// URL string exceeds Excel's limit of 2080 characters.
    MaxUrlLengthExceeded,
    /// Unknown URL type. The URL/URIs supported by Excel are `http://`,
    /// `https://`, `ftp://`, `ftps://`, `mailto:`, `file://`, and the
    /// pseudo-URI `internal:`.
    UnknownUrlType(String),
    /// Unknown image type. The supported image formats are PNG, JPG, GIF and
    /// BMP. See Image(crate::Image) for details.
    UnknownImageType,
    /// Image has zero width or height, or the dimensions couldn't be read.
    ImageDimensionError,
    /// A general error that is raised when a chart parameter is incorrect, or a
    /// chart is configured incorrectly.
    ChartError(String),
    /// A general error that is raised when a sparkline parameter is incorrect,
    /// or a sparkline is configured incorrectly.
    SparklineError(String),
    /// A general error when one of the parameters supplied to a
    /// ExcelDateTime(crate::ExcelDateTime) method is outside Excel's
    /// allowable ranges.
    ///
    /// Excel restricts dates to the range 1899-12-31 to 9999-12-31. For hours,
    /// the range is generally 0-24, although larger ranges can be used to
    /// indicate durations. Minutes should be in the range 0-60, and seconds
    /// should be in the range 0.0-59.999. Excel only supports millisecond
    /// resolution.
    DateTimeRangeError(String),
    /// A parsing error when trying to convert a string into an
    /// ExcelDateTime(crate::ExcelDateTime).
    ///
    /// The allowable date/time formats supported by
    /// ExcelDateTime.parseFromStr()(crate::ExcelDateTime::parse_from_str)
    /// are:
    ///
    /// ```text
    /// Dates:
    ///     yyyy-mm-dd
    ///
    /// Times:
    ///     hh:mm
    ///     hh:mm:ss
    ///     hh:mm:ss.sss
    ///
    /// DateTimes:
    ///     yyyy-mm-ddThh:mm:ss
    ///     yyyy-mm-dd hh:mm:ss
    /// ```
    ///
    /// The time part of `DateTimes` can contain optional or fractional seconds
    /// like the time examples. Timezone information is not supported by Excel
    /// and is ignored in the parsing.
    DateTimeParseError(String),
    /// The table range overlaps a previous table range. This is strictly
    /// prohibited by Excel.
    TableRangeOverlaps(String, String),
    /// A general error that is raised when a table parameter is incorrect, or a
    /// table is configured incorrectly.
    TableError(String),
    /// Table name is already in use in the workbook.
    TableNameReused(String),
    /// A Worksheet and Table autofilter range overlap. This is strictly
    /// prohibited by Excel.
    AutofilterRangeOverlaps(String, String),
    /// A general error that is raised when a conditional format parameter is
    /// incorrect or missing.
    ConditionalFormatError(String),
    /// A general error that is raised when a data validation parameter is
    /// incorrect or missing.
    DataValidationError(String),
    /// A general error raised when a VBA name doesn't meet Excel's criteria as
    /// defined by the following rules:
    ///
    /// - The name must be less than 32 characters.
    /// - The name can only contain word characters: letters, numbers and
    ///   underscores.
    /// - The name must start with a letter.
    /// - The name cannot be blank.
    VbaNameError(String),
    /// Excel limits the maximum worksheet group level to 8 levels.
    MaxGroupLevelExceeded,
    /// An error that is raised when setting the default format for a workbook
    /// or worksheet.
    DefaultFormatError(String),
    /// An error that is raised when setting the theme for a workbook.
    ThemeError(String),
    /// A customizable error that can be used by third parties to raise errors
    /// or as a conversion target for other error types.
    CustomError(String),
    /// Wrapper for a variety of std.io::Error errors such as file
    /// permissions when writing the xlsx file to disk. This can be caused by a
    /// non-existent parent directory or, commonly on Windows, if the file is
    /// already open in Excel.
    IoError(Error),
    /// Wrapper for a variety of zip.result::ZipError errors from
    /// zip.ZipWriter. These relate to errors arising from creating
    /// the xlsx file zip container.
    ZipError(ZipError),
}

impl From<XlsxError> for xlsx::XlsxError {
    fn from(value: XlsxError) -> xlsx::XlsxError {
        match value {
            XlsxError::ParameterError(v0) => xlsx::XlsxError::ParameterError(v0),
            XlsxError::RowColumnLimitError => xlsx::XlsxError::RowColumnLimitError,
            XlsxError::RowColumnOrderError => xlsx::XlsxError::RowColumnOrderError,
            XlsxError::SheetnameCannotBeBlank(v0) => xlsx::XlsxError::SheetnameCannotBeBlank(v0),
            XlsxError::SheetnameLengthExceeded(v0) => xlsx::XlsxError::SheetnameLengthExceeded(v0),
            XlsxError::SheetnameReused(v0) => xlsx::XlsxError::SheetnameReused(v0),
            XlsxError::SheetnameContainsInvalidCharacter(v0) => {
                xlsx::XlsxError::SheetnameContainsInvalidCharacter(v0)
            }
            XlsxError::SheetnameStartsOrEndsWithApostrophe(v0) => {
                xlsx::XlsxError::SheetnameStartsOrEndsWithApostrophe(v0)
            }
            XlsxError::MaxStringLengthExceeded => xlsx::XlsxError::MaxStringLengthExceeded,
            XlsxError::UnknownWorksheetNameOrIndex(v0) => {
                xlsx::XlsxError::UnknownWorksheetNameOrIndex(v0)
            }
            XlsxError::MergeRangeSingleCell => xlsx::XlsxError::MergeRangeSingleCell,
            XlsxError::MergeRangeOverlaps(v0, v1) => xlsx::XlsxError::MergeRangeOverlaps(v0, v1),
            XlsxError::MaxUrlLengthExceeded => xlsx::XlsxError::MaxUrlLengthExceeded,
            XlsxError::UnknownUrlType(v0) => xlsx::XlsxError::UnknownUrlType(v0),
            XlsxError::UnknownImageType => xlsx::XlsxError::UnknownImageType,
            XlsxError::ImageDimensionError => xlsx::XlsxError::ImageDimensionError,
            XlsxError::ChartError(v0) => xlsx::XlsxError::ChartError(v0),
            XlsxError::SparklineError(v0) => xlsx::XlsxError::SparklineError(v0),
            XlsxError::DateTimeRangeError(v0) => xlsx::XlsxError::DateTimeRangeError(v0),
            XlsxError::DateTimeParseError(v0) => xlsx::XlsxError::DateTimeParseError(v0),
            XlsxError::TableRangeOverlaps(v0, v1) => xlsx::XlsxError::TableRangeOverlaps(v0, v1),
            XlsxError::TableError(v0) => xlsx::XlsxError::TableError(v0),
            XlsxError::TableNameReused(v0) => xlsx::XlsxError::TableNameReused(v0),
            XlsxError::AutofilterRangeOverlaps(v0, v1) => {
                xlsx::XlsxError::AutofilterRangeOverlaps(v0, v1)
            }
            XlsxError::ConditionalFormatError(v0) => xlsx::XlsxError::ConditionalFormatError(v0),
            XlsxError::DataValidationError(v0) => xlsx::XlsxError::DataValidationError(v0),
            XlsxError::VbaNameError(v0) => xlsx::XlsxError::VbaNameError(v0),
            XlsxError::MaxGroupLevelExceeded => xlsx::XlsxError::MaxGroupLevelExceeded,
            XlsxError::DefaultFormatError(v0) => xlsx::XlsxError::DefaultFormatError(v0),
            XlsxError::ThemeError(v0) => xlsx::XlsxError::ThemeError(v0),
            XlsxError::CustomError(v0) => xlsx::XlsxError::CustomError(v0),
            XlsxError::IoError(v0) => xlsx::XlsxError::IoError(v0),
            XlsxError::ZipError(v0) => xlsx::XlsxError::ZipError(v0),
        }
    }
}
