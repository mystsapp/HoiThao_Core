using HoiThao_Core.Data.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using X.PagedList;

namespace HoiThao_Core.Data.Repository
{
    public interface IAseanRepository : IRepository<Asean>
    {
        void AddList(List<Asean> aseans);
    }

    public class AseanRepository : Repository<Asean>, IAseanRepository
    {
        public AseanRepository(hoinghiContext context) : base(context)
        {

        }

        public void AddList(List<Asean> aseans)
        {
            _context.AddRange(aseans);
            _context.SaveChanges();
        }
    }
}
