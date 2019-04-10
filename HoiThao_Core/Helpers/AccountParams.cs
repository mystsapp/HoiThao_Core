using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace HoiThao_Core.Helpers
{
    public class AccountParams
    {

        private const int MaxPagesize = 50;
        public int PageNumber { get; set; } = 1;
        private int pageSize = 10;
        public int PageSize
        {
            get { return pageSize; }
            set { pageSize = (value > MaxPagesize) ? MaxPagesize : value; }
        }
        public string searchString { get; set; }
    }
}
