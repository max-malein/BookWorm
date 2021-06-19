﻿// <copyright file="ConstructCellBorder.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace BookWorm.Construct
{
    using System;
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
            pManager.AddIntegerParameter("Style", "S", "0 - STYLE_UNSPECIFIED, 1  -DOTTED, 2 - DASHED, 3 - SOLID, 4 - SOLID_MEDIUM, 5 - SOLID_THICK, 6 - NONE, 7 - DOUBLE", GH_ParamAccess.item);
            pManager.AddColourParameter("Color", "C", "Border color", GH_ParamAccess.item);
        }

        /// <inheritdoc/>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Border", "B", "Border", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var style = Style.NONE;
            var color = new Color();

            if (!DA.GetData(0, ref style))
            {
                return;
            }

            if (!DA.GetData(1, ref color))
            {
                return;
            }

            var border = new Border()
            {
                Style = style.ToString(),
                Color = color,
            };

            DA.SetData(0, border);
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