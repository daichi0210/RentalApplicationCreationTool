using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RentalApplicationCreationTool
{
    internal class AddToList
    {
        public string TextFormatting(string list, string contents) {
            string text;

            if (list != "")
            {
                text = "、" + contents;
            }
            else
            {
                text = contents;
            }

            return text;
        }
    }
}
