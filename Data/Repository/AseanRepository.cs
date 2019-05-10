using Data.Interfaces;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using X.PagedList;

namespace Data.Repository
{
    public interface IAseanRepository : IRepository<Asean>
    {
        void AddList(List<Asean> aseans);

        IPagedList<Asean> GetAseans(string option, string searchString, int? page);

        DataTable ConferenceReport();

        DataTable ConferenceGroupByCountry();

        IEnumerable<string> GetAllCountry();

        DataTable GetAllByCountry(string countryName);

        DataTable PaymentReport();

        DataTable PickupReport();

        DataTable AirticketReport();

        DataTable TourReport();

        IEnumerable<string> GetAllHotel();

        DataTable HotelReport(string hotel);

        DataTable CheckinReport();

        DataTable VatReport();

        DataTable SpeakerReport();
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

        public DataTable ConferenceReport()
        {
            try
            {
                DataTable dt = new DataTable();
                var result = _context.Aseans.Select(p => new
                {
                    p.K,
                    p.Dangky,
                    p.Firstname,
                    p.Address,
                    p.Country,
                    p.Tel,
                    p.Email,
                    p.Id,
                    p.Totala,
                    p.Totalb,
                    p.Grandtotal,
                    p.Payment,
                    p.Currency,
                    p.Lastname,
                    p.Title,
                    p.Cardnumber
                });
                var count = result.Count();

                dt = EntityToTable.ToDataTable(result);
                if (dt.Rows.Count > 0)
                    return dt;
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public DataTable ConferenceGroupByCountry()
        {
            try
            {
                DataTable dt = new DataTable();
                var result = _context.Aseans.Select(p => new
                {
                    p.K,
                    p.Dangky,
                    p.Firstname,
                    p.Address,
                    p.Country,
                    p.Tel,
                    p.Email,
                    p.Id,
                    p.Totala,
                    p.Totalb,
                    p.Grandtotal,
                    p.Payment,
                    p.Currency,
                    p.Lastname,
                    p.Title,
                    p.Cardnumber
                }).OrderBy(x => x.Country);

                var count = result.Count();

                dt = EntityToTable.ToDataTable(result);
                if (dt.Rows.Count > 0)
                    return dt;
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public IEnumerable<string> GetAllCountry()
        {
            return _context.Aseans.Select(x => x.Country).Distinct();
        }

        public DataTable GetAllByCountry(string countryName)
        {
            var result = _context.Aseans.Where(x => x.Country == countryName).Select(p => new
            {
                p.K,
                p.Dangky,
                p.Firstname,
                p.Address,
                p.Country,
                p.Tel,
                p.Email,
                p.Id,
                p.Totala,
                p.Totalb,
                p.Grandtotal,
                p.Payment,
                p.Currency,
                p.Lastname,
                p.Title,
                p.Cardnumber
            });
            DataTable dt = new DataTable();
            dt = EntityToTable.ToDataTable(result);
            if (dt.Rows.Count > 0)
                return dt;
            else
                return null;
        }

        public DataTable PaymentReport()
        {
            try
            {
                DataTable dt = new DataTable();
                var result = _context.Aseans.Select(p => new
                {
                    p.K,
                    p.Dangky,
                    p.Firstname,
                    p.Address,
                    p.Department,
                    p.Tel,
                    p.Email,
                    p.Id,
                    p.Totala,
                    p.Totalb,
                    p.Grandtotal,
                    p.Payment,
                    p.Currency,
                    p.Lastname,
                    p.Title,
                    p.Cardnumber,
                    p.Invited,
                    p.Amount,
                    p.Country,
                    p.Cabk,
                    p.Caravelle,
                    p.CarAh,
                    p.Hotel,
                    p.Note
                });

                var count = result.Count();

                dt = EntityToTable.ToDataTable(result);
                if (dt.Rows.Count > 0)
                    return dt;
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public DataTable PickupReport()
        {
            try
            {
                DataTable dt = new DataTable();
                var result = _context.Aseans.Select(p => new
                {
                    p.K,
                    p.Dangky,
                    p.Firstname,
                    p.Address,
                    p.Department,
                    p.Tel,
                    p.Email,
                    p.Id,
                    p.Totala,
                    p.Totalb,
                    p.Grandtotal,
                    p.Payment,
                    p.Currency,
                    p.Lastname,
                    p.Title,
                    p.Cardnumber,
                    p.Invited,
                    p.Amount
                });

                var count = result.Count();

                dt = EntityToTable.ToDataTable(result);
                if (dt.Rows.Count > 0)
                    return dt;
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public DataTable AirticketReport()
        {
            try
            {
                DataTable dt = new DataTable();
                var result = _context.Aseans.Select(p => new
                {
                    p.K,
                    p.Dangky,
                    p.Firstname,
                    p.Address,
                    p.Department,
                    p.Tel,
                    p.Email,
                    p.Id,
                    p.Ydepdate,
                    p.Cdepdate
                });

                var count = result.Count();

                dt = EntityToTable.ToDataTable(result);
                if (dt.Rows.Count > 0)
                    return dt;
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public DataTable TourReport()
        {
            try
            {
                DataTable dt = new DataTable();
                var result = _context.Aseans.Select(p => new
                {
                    p.K,
                    p.Dangky,
                    p.Firstname,
                    p.Address,
                    p.Department,
                    p.Tel,
                    p.Email,
                    p.Id,
                    p.Fax,
                    p.Hotel
                });

                var count = result.Count();

                dt = EntityToTable.ToDataTable(result);
                if (dt.Rows.Count > 0)
                    return dt;
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public IEnumerable<string> GetAllHotel()
        {
            return _context.Aseans.Select(x => x.Hotel).Distinct();
        }

        public DataTable HotelReport(string hotel)
        {
            DataTable dt = new DataTable();
            if (hotel == "Other")
            {
                var result = _context.Aseans.Where(x => x.Hotel == null).Select(p => new
                {
                    p.K,
                    p.Dangky,
                    p.Firstname,
                    p.Address,
                    p.Department,
                    p.Tel,
                    p.Email,
                    p.Id,
                    p.HotelCheckin,
                    p.HotelCheckout,
                    p.HotelPrice,
                    p.HotelBookingInf,
                    p.Group
                });
                var count = result.Count();
                dt = EntityToTable.ToDataTable(result);
            }
            else
            {
                var result = _context.Aseans.Where(x => x.Hotel == hotel).Select(p => new
                {
                    p.K,
                    p.Dangky,
                    p.Firstname,
                    p.Address,
                    p.Department,
                    p.Tel,
                    p.Email,
                    p.Id,
                    p.HotelCheckin,
                    p.HotelCheckout,
                    p.HotelPrice,
                    p.HotelBookingInf,
                    p.Group
                });
                var count = result.Count();
                dt = EntityToTable.ToDataTable(result);
            }

            if (dt.Rows.Count > 0)
                return dt;
            else
                return null;
        }

        public DataTable CheckinReport()
        {
            try
            {
                DataTable dt = new DataTable();
                var result = _context.Aseans.Where(x => x.Checkin.HasValue).Select(p => new
                {
                    p.K,
                    p.Checkin,
                    p.Firstname,
                    p.Address,
                    p.Country,
                    p.Tel,
                    p.Email,
                    p.Id,
                    p.Totala,
                    p.Totalb,
                    p.Grandtotal,
                    p.Payment,
                    p.Currency,
                    p.Lastname,
                    p.Title,
                    p.Cardnumber
                });
                var count = result.Count();

                dt = EntityToTable.ToDataTable(result);
                if (dt.Rows.Count > 0)
                    return dt;
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public DataTable VatReport()
        {
            try
            {
                DataTable dt = new DataTable();
                var result = _context.Aseans.Select(p => new
                {
                    p.K,
                    p.Id,
                    p.Firstname,
                    p.Company,
                    p.Descript,
                    p.Dangky,
                    p.Currency,
                    p.Vatbill,
                    p.Rate,
                    p.Totalb,
                    p.Bankfee,
                    p.Payment,
                    p.Grandtotal,
                    p.Taxcode
                });
                var count = result.Count();

                dt = EntityToTable.ToDataTable(result);
                if (dt.Rows.Count > 0)
                    return dt;
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }

        public DataTable SpeakerReport()
        {
            try
            {
                DataTable dt = new DataTable();
                var result = _context.Aseans.Where(x => x.Checkin.HasValue && (x.Speaker == true)).Select(p => new
                {
                    p.K,
                    p.Checkin,
                    p.Firstname,
                    p.Address,
                    p.Country,
                    p.Tel,
                    p.Email,
                    p.Id,
                    p.Totala,
                    p.Totalb,
                    p.Grandtotal,
                    p.Payment,
                    p.Currency,
                    p.Lastname,
                    p.Title,
                    p.Cardnumber
                });
                var count = result.Count();

                dt = EntityToTable.ToDataTable(result);
                if (dt.Rows.Count > 0)
                    return dt;
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }
    }
}