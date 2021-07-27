using System;
using System.Drawing;
using Grasshopper.Kernel;

namespace GoogleDocs
{
    /// <summary>
    /// Info.
    /// </summary>
    public class BookWormInfo : GH_AssemblyInfo
    {
        /// <inheritdoc/>
        public override string Name
        {
            get
            {
                return "Bookworm";
            }
        }

        /// <inheritdoc/>
        public override Bitmap Icon
        {
            get
            {
                // Return a 24x24 pixel bitmap to represent this GHA library.
                return null;
            }
        }

        /// <inheritdoc/>
        public override string Description => "GoogleDocs connector. The description is really long and thorough";

        /// <inheritdoc/>
        public override string Version => "0.9.1";

        /// <inheritdoc/>
        public override Guid Id
        {
            get
            {
                return new Guid("56dfe1a3-4e7b-425f-b169-965c0d1f7977");
            }
        }

        /// <inheritdoc/>
        public override string AuthorName
        {
            get
            {
                // Return a string identifying you or your company.
                return "Max Malein";
            }
        }

        /// <inheritdoc/>
        public override string AuthorContact
        {
            get
            {
                // Return a string representing your preferred contact details.
                return @"https://github.com/max-malein/BookWorm";
            }
        }
    }
}
