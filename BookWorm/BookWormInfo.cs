using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace GoogleDocs
{
    public class BookWormInfo : GH_AssemblyInfo
    {
        public override string Name
        {
            get
            {
                return "Bookworm";
            }
        }

        public override Bitmap Icon
        {
            get
            {
                //Return a 24x24 pixel bitmap to represent this GHA library.
                return null;
            }
        }

        public override string Description => "GoogleDocs connector. The description is really long and thorough";
        public override string Version => "0.6.2";

        public override Guid Id
        {
            get
            {
                return new Guid("56dfe1a3-4e7b-425f-b169-965c0d1f7977");
            }
        }

        public override string AuthorName
        {
            get
            {
                //Return a string identifying you or your company.
                return "Max Malein";
            }
        }

        public override string AuthorContact
        {
            get
            {
                //Return a string representing your preferred contact details.
                return @"https://github.com/max-malein/BookWorm";
            }
        }
    }
}
