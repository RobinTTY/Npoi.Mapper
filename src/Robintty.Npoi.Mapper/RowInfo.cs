﻿namespace Robintty.Npoi.Mapper
{
    /// <summary>
    /// Information about a single row of an Excel worksheet.
    /// </summary>
    /// <typeparam name="TTarget">The target mapping type for a row.</typeparam>
    public class RowInfo<TTarget> : IRowInfo
    {
        /// <summary>
        /// The row number.
        /// </summary>
        public int RowNumber { get; set; }

        // TODO: what values can this have?
        /// <summary>
        /// Constructed value object from the row.
        /// </summary>
        public TTarget Value { get;  set; }

        /// <summary>
        /// Column index of the first error.
        /// </summary>
        public int ErrorColumnIndex { get; set; }

        /// <summary>
        /// Error message of the first error.
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Creates a new RowData object.
        /// </summary>
        /// <param name="rowNumber">The row number.</param>
        /// <param name="value">Constructed value object from the row.</param>
        /// <param name="errorColumnIndex">Column index of the first error.</param>
        /// <param name="errorMessage">Error message of the first error.</param>
        public RowInfo(int rowNumber, TTarget value, int errorColumnIndex, string errorMessage)
        {
            RowNumber = rowNumber;
            Value = value;
            ErrorColumnIndex = errorColumnIndex;
            ErrorMessage = errorMessage;
        }
    }
}
