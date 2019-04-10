using HoiThao_Core.Data.Interfaces;
using HoiThao_Core.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using X.PagedList;

namespace HoiThao_Core.Data.Repository
{
    public interface IAccountRepository : IRepository<Account>
    {
        IPagedList<Account> GetAccounts(string searchString, int? page);
    }

    public class AccountRepository : Repository<Account>, IAccountRepository
    {
        public AccountRepository(hoinghiContext context) : base(context)
        {

        }

        public IPagedList<Account> GetAccounts(string searchString, int? page)
        {
            // return a 404 if user browses to before the first page
            if (page.HasValue && page < 1)
                return null;

            // retrieve list from database/whereverand
            var listAccount = _context.Accounts.AsQueryable();

            if (!string.IsNullOrEmpty(searchString))
                listAccount = listAccount.Where(a => a.Username.Contains(searchString) || a.Hoten.Contains(searchString));

            // page the list
            const int pageSize = 1;
            var listPaged = listAccount.ToPagedList(page ?? 1, pageSize);

            // return a 404 if user browses to pages beyond last page. special case first page if no items exist
            if (listPaged.PageNumber != 1 && page.HasValue && page > listPaged.PageCount)
                return null;

            return listPaged;
        }
    }
}
