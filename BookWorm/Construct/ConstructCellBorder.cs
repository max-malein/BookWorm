// <copyright file="ConstructCellBorder.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace BookWorm.Construct
{
    using System;
    using BookWorm.Goo;
    using Google.Apis.Sheets.v4.Data;
    using Grasshopper.Kernel;

    /// <summary>
    /// Construct cell border.
    /// </summary>
    public class ConstructCellBorder : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConstructCellBorder"/> class.
        /// </summary>
        public ConstructCellBorder()
           : base(
                 "ConstructCellBorder",
                 "ConCellBorder",
                 "Construct cell border",
                 "BookWorm",
                 "Construct")
        {
        }

        /// <inheritdoc/>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("Style", "Style", "0 - STYLE_UNSPECIFIED, 1  -DOTTED, 2 - DASHED, 3 - SOLID, 4 - SOLID_MEDIUM, 5 - SOLID_THICK, 6 - NONE, 7 - DOUBLE", GH_ParamAccess.item);
            pManager.AddTextParameter("Color", "Color", "Border color", GH_ParamAccess.item);

            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Border", "Border", "Border", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            
            var style = 0;
            var color = string.Empty;
            var border = new Border();

            if (DA.GetData(0, ref style))
            {
                border.Style = style.ToString();
            }

            if (DA.GetData(1, ref color))
            {
                border.Color = new Color
                {
                    Alpha = (float)255,
                    Red = float.Parse(color.Split(',')[0]),
                    Green = float.Parse(color.Split(',')[1]),
                    Blue = float.Parse(color.Split(',')[2]),
                };
            }

            

            var borderGoo = new GH_CellBorder(border);
            DA.SetData(0, borderGoo);
        }

        /// <summary>
        /// Gets provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                // You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid => new Guid("30b79fc6-f04c-4aac-9aaf-80a994e303e0");

        /// <summary>
        /// DOTTED The border is dotted.
        /// DASHED The border is dashed.
        /// SOLID The border is a thin solid line.
        /// SOLID_MEDIUM    The border is a medium solid line.
        /// SOLID_THICK The border is a thick solid line.
        /// NONE No border.Used only when updating a border in order to erase it.
        /// DOUBLE The border is two solid lines.
        /// </summary>
        private enum Style
        {
            STYLE_UNSPECIFIED,
            DOTTED,
            DASHED,
            SOLID,
            SOLID_MEDIUM,
            SOLID_THICK,
            NONE,
            DOUBLE,
        }
    }
}