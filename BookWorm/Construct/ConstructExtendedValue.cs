using BookWorm.Goo;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;

namespace BookWorm.Construct
{
    public class ConstructExtendedValue : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ConstructExtendedValue class.
        /// </summary>
        public ConstructExtendedValue()
          : base(
                "ConstructExtendedValue",
                "ConExtendedValue",
                 "Construct extended value",
                 "BookWorm",
                 "Construct")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Value", "Value", "Value", GH_ParamAccess.item);
           
            for (int i = 0; i < pManager.ParamCount; i++)
            {
                pManager[i].Optional = true;
            }
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("ExtendedValue", "ExtendedValue", "Extended value", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var ext = new ExtendedValue();
            var enterString = string.Empty;

            DA.GetData(0, ref enterString);

            if(enterString==string.Empty)
            {
                ext.StringValue = enterString;
            }
            else if(enterString.ToLower()=="false")
            {
                ext.BoolValue = false;
            }
            else if (enterString.ToLower() == "true")
            {
                ext.BoolValue = true;
            } else if (double.TryParse(enterString, out double number))
            {
                ext.NumberValue = number;
            }else if(enterString.ToCharArray()[0]=='=')
            {
                ext.FormulaValue = enterString;
            }else 
            {
                ext.StringValue = enterString;
            }
 
            var extGoo = new GH_ExtendedValue(ext);
            DA.SetData(0, extGoo);
        }

        /// <summary>
        /// Provides an Icon for the component.
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
        public override Guid ComponentGuid
        {
            get { return new Guid("36d6cbe5-83f3-4d7b-941d-a16a525da60a"); }
        }
    }
}