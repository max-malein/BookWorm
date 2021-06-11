using System;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel.Types;
using Newtonsoft.Json;

namespace BookWorm.Goo
{
    /// <summary>
    /// ExtendedValue Goo.
    /// </summary>
    public class GH_ExtendedValue : GH_Goo<ExtendedValue>
    {
        /// <summary>
        /// Default.
        /// </summary>
        public GH_ExtendedValue()
        {
        }

        /// <summary>
        /// ExtendedValue Goo.
        /// </summary>
        /// <param name="extendedValue">ExtendedValue.</param>
        public GH_ExtendedValue(ExtendedValue extendedValue)
        {
            Value = extendedValue;
        }

        /// <summary>
        /// Deep copy of Goo.
        /// </summary>
        /// <param name="extendedValueGoo">ExtendedValue Goo.</param>
        public GH_ExtendedValue(GH_ExtendedValue extendedValueGoo)
        {
            if (extendedValueGoo != null)
            {
                // смекалОчка - у ячейки нет своего копирования.
                var extendedValueJson = JsonConvert.SerializeObject(extendedValueGoo.Value, Formatting.Indented);
                var extendedValue = JsonConvert.DeserializeObject<ExtendedValue>(extendedValueJson);
                Value = extendedValue;
            }
        }

        /// <inheritdoc/>
        public override bool IsValid => Value != null;

        /// <inheritdoc/>
        public override string TypeName => "ExtendedValue";

        /// <inheritdoc/>
        public override string TypeDescription => "The kinds of value that a cell in a spreadsheet can have";

        /// <inheritdoc/>
        public override IGH_Goo Duplicate()
        {
            return new GH_ExtendedValue(this);
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            if (Value == null)
            {
                return string.Empty;
            }

            var extendedValueAsString = string.Empty;

            if (Value.NumberValue != null)
            {
                extendedValueAsString = $@"User double: {Value.NumberValue}";
            }
            else if (Value.StringValue != null)
            {
                extendedValueAsString = $@"User string: {Value.StringValue}";
            }
            else if (Value.BoolValue != null)
            {
                extendedValueAsString = $@"User bool: {Value.BoolValue}";
            }
            else if (Value.FormulaValue != null)
            {
                extendedValueAsString = $@"User formula: {Value.FormulaValue}";
            }
            else if (Value.ErrorValue != null)
            {
                extendedValueAsString = $@"Error: {Value.ErrorValue.Type}. {Value.ErrorValue.Message}";
            }

            return extendedValueAsString;
        }
    }
}
