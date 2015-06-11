using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToJsonApp
{
   public class ViewJsonResult
    {
        public string Index
        {
            get;
            set;
        }

        public List<ViewValue> Value
        {
            get;
            set;
        }
    }
}
