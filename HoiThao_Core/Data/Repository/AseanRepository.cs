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
        IPagedList<Asean> GetAseans(string option, string searchString, int? page);
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

        public IPagedList<Asean> GetAseans(string option, string searchString, int? page)
        {
            // return a 404 if user browses to before the first page
            if (page.HasValue && page < 1)
                return null;

            // retrieve list from database/whereverand
            var list = _context.Aseans.AsQueryable();
           
            if (!string.IsNullOrEmpty(option))
            {
                bool statusBool = bool.Parse(option);

                if (statusBool)
                {
                    // retrieve list from database/whereverand
                    list = _context.Aseans.Where(i => i.Invited == true).AsQueryable();
                }
                if (!statusBool)
                {
                    // retrieve list from database/whereverand
                    list = _context.Aseans.Where(i => i.Speaker == true).AsQueryable();
                }
            }
            

            if (!string.IsNullOrEmpty(searchString))
                list = list.Where(a => a.Firstname.Contains(searchString) || a.Id.Contains(searchString));

            // page the list
            const int pageSize = 5;
            var listPaged = list.ToPagedList(page ?? 1, pageSize);

            // return a 404 if user browses to pages beyond last page. special case first page if no items exist
            if (listPaged.PageNumber != 1 && page.HasValue && page > listPaged.PageCount)
                return null;

            return listPaged;
        }
    }
}
