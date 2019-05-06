using HoiThao_Core.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace HoiThao_Core.Models
{
    public class EditAseanVM
    {
        public Asean Asean { get; set; }

        public List<OptionListVM> OptionListVM { get; set; }
    }
}
