using Microsoft.EntityFrameworkCore;

namespace Data
{
    public partial class hoinghiContext : DbContext
    {
        public hoinghiContext()
        {
        }

        public hoinghiContext(DbContextOptions<hoinghiContext> options)
            : base(options)
        {
        }

        public virtual DbSet<Account> Accounts { get; set; }
        public virtual DbSet<Asean> Aseans { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlServer("Server=192.168.4.198;Database=hoinghi;Trusted_Connection=False; User Id=sa; Password=TigerSts@2017");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.HasAnnotation("ProductVersion", "2.2.3-servicing-35854");

            modelBuilder.Entity<Account>(entity =>
            {
                entity.ToTable("account");

                entity.Property(e => e.Id).HasColumnName("id");

                entity.Property(e => e.Chinhanh)
                    .HasColumnName("chinhanh")
                    .HasMaxLength(3)
                    .IsUnicode(false);

                entity.Property(e => e.Daily)
                    .HasColumnName("daily")
                    .HasMaxLength(50);

                entity.Property(e => e.Doimatkhau).HasColumnName("doimatkhau");

                entity.Property(e => e.Hoten)
                    .HasColumnName("hoten")
                    .HasMaxLength(50);

                entity.Property(e => e.Khoi)
                    .HasColumnName("khoi")
                    .HasMaxLength(5)
                    .IsUnicode(false)
                    .HasDefaultValueSql("('OB')");

                entity.Property(e => e.Ngaycapnhat)
                    .HasColumnName("ngaycapnhat")
                    .HasColumnType("datetime");

                entity.Property(e => e.Ngaydoimk)
                    .HasColumnName("ngaydoimk")
                    .HasColumnType("date")
                    .HasDefaultValueSql("(getdate())");

                entity.Property(e => e.Ngaytao)
                    .HasColumnName("ngaytao")
                    .HasColumnType("datetime")
                    .HasDefaultValueSql("(getdate())");

                entity.Property(e => e.Nguoicapnhat)
                    .HasColumnName("nguoicapnhat")
                    .HasMaxLength(50);

                entity.Property(e => e.Nguoitao)
                    .HasColumnName("nguoitao")
                    .HasMaxLength(50);

                entity.Property(e => e.Nhom)
                    .HasColumnName("nhom")
                    .HasMaxLength(50)
                    .IsUnicode(false);

                entity.Property(e => e.Password)
                    .HasColumnName("password")
                    .HasMaxLength(50);

                entity.Property(e => e.Role)
                    .HasColumnName("role")
                    .HasMaxLength(50)
                    .IsUnicode(false);

                entity.Property(e => e.Trangthai)
                    .IsRequired()
                    .HasColumnName("trangthai")
                    .HasDefaultValueSql("((1))");

                entity.Property(e => e.Username)
                    .IsRequired()
                    .HasColumnName("username")
                    .HasMaxLength(50);
            });

            modelBuilder.Entity<Asean>(entity =>
            {
                entity.HasKey(e => e.K);

                entity.ToTable("asean");

                entity.Property(e => e.K)
                    .HasColumnName("k")
                    .HasColumnType("decimal(18, 0)")
                    .ValueGeneratedOnAdd();

                entity.Property(e => e.Address)
                    .HasColumnName("address")
                    .HasMaxLength(200);

                entity.Property(e => e.Afno)
                    .HasColumnName("afno")
                    .HasMaxLength(120);

                entity.Property(e => e.Allergy)
                    .HasColumnName("allergy")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Amount)
                    .HasColumnName("amount")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.At)
                    .HasColumnName("at")
                    .HasMaxLength(20);

                entity.Property(e => e.Bankfee)
                    .HasColumnName("bankfee")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Bongsen).HasColumnName("bongsen");

                entity.Property(e => e.BsRate)
                    .HasColumnName("bs_rate")
                    .HasMaxLength(40);

                entity.Property(e => e.Bsbk)
                    .HasColumnName("bsbk")
                    .HasMaxLength(50);

                entity.Property(e => e.Bschinh)
                    .HasColumnName("bschinh")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Bsnoitru)
                    .HasColumnName("bsnoitru")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.BusAh)
                    .HasColumnName("bus_ah")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.BusHa)
                    .HasColumnName("bus_ha")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.CaRate)
                    .HasColumnName("ca_rate")
                    .HasMaxLength(40);

                entity.Property(e => e.Cabk)
                    .HasColumnName("cabk")
                    .HasMaxLength(50);

                entity.Property(e => e.CarAh)
                    .HasColumnName("car_ah")
                    .HasMaxLength(10);

                entity.Property(e => e.CarHa)
                    .HasColumnName("car_ha")
                    .HasMaxLength(10);

                entity.Property(e => e.Caravelle).HasColumnName("caravelle");

                entity.Property(e => e.Cardnumber)
                    .HasColumnName("cardnumber")
                    .HasMaxLength(20);

                entity.Property(e => e.Cdepdate)
                    .HasColumnName("cdepdate")
                    .HasMaxLength(10);

                entity.Property(e => e.Cdepperson)
                    .HasColumnName("cdepperson")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Cdeproute)
                    .HasColumnName("cdeproute")
                    .HasMaxLength(10);

                entity.Property(e => e.Central)
                    .HasColumnName("central")
                    .HasMaxLength(10);

                entity.Property(e => e.Checkin)
                    .HasColumnName("checkin")
                    .HasColumnType("datetime");

                entity.Property(e => e.Cibongsen)
                    .HasColumnName("cibongsen")
                    .HasMaxLength(10);

                entity.Property(e => e.Cicaravel)
                    .HasColumnName("cicaravel")
                    .HasMaxLength(10);

                entity.Property(e => e.Cikimdo)
                    .HasColumnName("cikimdo")
                    .HasMaxLength(10);

                entity.Property(e => e.Cirex)
                    .HasColumnName("cirex")
                    .HasMaxLength(10);

                entity.Property(e => e.Cisheraton)
                    .HasColumnName("cisheraton")
                    .HasMaxLength(10);

                entity.Property(e => e.Cisofitel)
                    .HasColumnName("cisofitel")
                    .HasMaxLength(10);

                entity.Property(e => e.City)
                    .HasColumnName("city")
                    .HasColumnType("datetime");

                entity.Property(e => e.Cityper)
                    .HasColumnName("cityper")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Ciwinsor)
                    .HasColumnName("ciwinsor")
                    .HasMaxLength(10);

                entity.Property(e => e.Cobongsen)
                    .HasColumnName("cobongsen")
                    .HasMaxLength(10);

                entity.Property(e => e.Cocaravel)
                    .HasColumnName("cocaravel")
                    .HasMaxLength(10);

                entity.Property(e => e.Code).HasColumnName("code");

                entity.Property(e => e.Cokimdo)
                    .HasColumnName("cokimdo")
                    .HasMaxLength(10);

                entity.Property(e => e.Company)
                    .HasColumnName("company")
                    .HasMaxLength(150);

                entity.Property(e => e.Corex)
                    .HasColumnName("corex")
                    .HasMaxLength(10);

                entity.Property(e => e.Cosheraton)
                    .HasColumnName("cosheraton")
                    .HasMaxLength(10);

                entity.Property(e => e.Cosofitel)
                    .HasColumnName("cosofitel")
                    .HasMaxLength(10);

                entity.Property(e => e.Country)
                    .HasColumnName("country")
                    .HasMaxLength(50);

                entity.Property(e => e.Cowinsor)
                    .HasColumnName("cowinsor")
                    .HasMaxLength(10);

                entity.Property(e => e.Cretdate)
                    .HasColumnName("cretdate")
                    .HasMaxLength(10);

                entity.Property(e => e.Cretperson)
                    .HasColumnName("cretperson")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Cretroute)
                    .HasColumnName("cretroute")
                    .HasMaxLength(10);

                entity.Property(e => e.Ctper)
                    .HasColumnName("ctper")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Cuchi)
                    .HasColumnName("cuchi")
                    .HasColumnType("datetime");

                entity.Property(e => e.Cuchiper)
                    .HasColumnName("cuchiper")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Currency)
                    .HasColumnName("currency")
                    .HasMaxLength(10);

                entity.Property(e => e.Dangky)
                    .HasColumnName("dangky")
                    .HasColumnType("datetime");

                entity.Property(e => e.Deepmk)
                    .HasColumnName("deepmk")
                    .HasColumnType("datetime");

                entity.Property(e => e.Deepmkper)
                    .HasColumnName("deepmkper")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Department)
                    .HasColumnName("department")
                    .HasMaxLength(200);

                entity.Property(e => e.Descript)
                    .HasColumnName("descript")
                    .HasMaxLength(150);

                entity.Property(e => e.Dfno)
                    .HasColumnName("dfno")
                    .HasMaxLength(120);

                entity.Property(e => e.Dikem)
                    .HasColumnName("dikem")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Dt)
                    .HasColumnName("dt")
                    .HasMaxLength(120);

                entity.Property(e => e.Email)
                    .HasColumnName("email")
                    .HasMaxLength(150);

                entity.Property(e => e.Fax)
                    .HasColumnName("fax")
                    .HasMaxLength(30);

                entity.Property(e => e.Firstname)
                    .HasColumnName("firstname")
                    .HasMaxLength(30);

                entity.Property(e => e.Giaallergy)
                    .HasColumnName("giaallergy")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Giachinh)
                    .HasColumnName("giachinh")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Giadikem)
                    .HasColumnName("giadikem")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Giaht1)
                    .HasColumnName("giaht1")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Giaht2)
                    .HasColumnName("giaht2")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Giaht3)
                    .HasColumnName("giaht3")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Giahtgp1)
                    .HasColumnName("giahtgp1")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Giahtgp2)
                    .HasColumnName("giahtgp2")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Giahtgp3)
                    .HasColumnName("giahtgp3")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Gianoitru)
                    .HasColumnName("gianoitru")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Golf)
                    .HasColumnName("golf")
                    .HasColumnType("datetime");

                entity.Property(e => e.Golfper)
                    .HasColumnName("golfper")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Grandtotal)
                    .HasColumnName("grandtotal")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Group)
                    .HasColumnName("group")
                    .HasMaxLength(200);

                entity.Property(e => e.Hoithao1)
                    .HasColumnName("hoithao1")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Hoithao2)
                    .HasColumnName("hoithao2")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Hoithao3)
                    .HasColumnName("hoithao3")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Hoithaogp1)
                    .HasColumnName("hoithaogp1")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Hoithaogp2)
                    .HasColumnName("hoithaogp2")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Hoithaogp3)
                    .HasColumnName("hoithaogp3")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Hotel).HasMaxLength(150);

                entity.Property(e => e.HotelBookingInf).HasMaxLength(250);

                entity.Property(e => e.HotelCheckin).HasMaxLength(50);

                entity.Property(e => e.HotelCheckout).HasMaxLength(50);

                entity.Property(e => e.HotelDon).HasMaxLength(50);

                entity.Property(e => e.HotelPrice).HasMaxLength(50);

                entity.Property(e => e.HotelTien).HasMaxLength(50);

                entity.Property(e => e.Htlother)
                    .HasColumnName("htlother")
                    .HasMaxLength(50);

                entity.Property(e => e.Id)
                    .HasColumnName("id")
                    .HasMaxLength(10);

                entity.Property(e => e.Institutio)
                    .HasColumnName("institutio")
                    .HasMaxLength(200);

                entity.Property(e => e.Invited).HasColumnName("invited");

                entity.Property(e => e.KdRate)
                    .HasColumnName("kd_rate")
                    .HasMaxLength(40);

                entity.Property(e => e.Kdbk)
                    .HasColumnName("kdbk")
                    .HasMaxLength(50);

                entity.Property(e => e.Kimdo).HasColumnName("kimdo");

                entity.Property(e => e.Kpc)
                    .HasColumnName("kpc")
                    .HasColumnType("datetime");

                entity.Property(e => e.Kpcsk)
                    .HasColumnName("kpcsk")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Lao)
                    .HasColumnName("lao")
                    .HasColumnType("datetime");

                entity.Property(e => e.Laosk)
                    .HasColumnName("laosk")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Lastname)
                    .HasColumnName("lastname")
                    .HasMaxLength(50);

                entity.Property(e => e.Makh)
                    .HasColumnName("makh")
                    .HasMaxLength(10);

                entity.Property(e => e.Mekong)
                    .HasColumnName("mekong")
                    .HasColumnType("datetime");

                entity.Property(e => e.Mekongper)
                    .HasColumnName("mekongper")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Mop)
                    .HasColumnName("mop")
                    .HasMaxLength(50);

                entity.Property(e => e.Nhatrang)
                    .HasColumnName("nhatrang")
                    .HasColumnType("datetime");

                entity.Property(e => e.Nhatrangsk)
                    .HasColumnName("nhatrangsk")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Note).HasColumnName("note");

                entity.Property(e => e.Payment)
                    .HasColumnName("payment")
                    .HasMaxLength(50);

                entity.Property(e => e.Phanthiet)
                    .HasColumnName("phanthiet")
                    .HasColumnType("datetime");

                entity.Property(e => e.Ptper)
                    .HasColumnName("ptper")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Rate)
                    .HasColumnName("rate")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Rex).HasColumnName("rex");

                entity.Property(e => e.RexRate)
                    .HasColumnName("rex_rate")
                    .HasMaxLength(40);

                entity.Property(e => e.Rexbk)
                    .HasColumnName("rexbk")
                    .HasMaxLength(50);

                entity.Property(e => e.Saigon)
                    .HasColumnName("saigon")
                    .HasColumnType("datetime");

                entity.Property(e => e.Saigonper)
                    .HasColumnName("saigonper")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.SfRate)
                    .HasColumnName("sf_rate")
                    .HasMaxLength(40);

                entity.Property(e => e.Sfbk)
                    .HasColumnName("sfbk")
                    .HasMaxLength(50);

                entity.Property(e => e.ShRate)
                    .HasColumnName("sh_rate")
                    .HasMaxLength(40);

                entity.Property(e => e.Sharecarah)
                    .HasColumnName("sharecarah")
                    .HasMaxLength(60);

                entity.Property(e => e.Sharecarha)
                    .HasColumnName("sharecarha")
                    .HasMaxLength(60);

                entity.Property(e => e.Shbk)
                    .HasColumnName("shbk")
                    .HasMaxLength(50);

                entity.Property(e => e.Sheraton).HasColumnName("sheraton");

                entity.Property(e => e.Sofitel).HasColumnName("sofitel");

                entity.Property(e => e.Speaker).HasColumnName("speaker");

                entity.Property(e => e.Taxcode)
                    .HasColumnName("taxcode")
                    .HasMaxLength(16);

                entity.Property(e => e.Tel)
                    .HasColumnName("tel")
                    .HasMaxLength(30);

                entity.Property(e => e.Title)
                    .HasColumnName("title")
                    .HasMaxLength(30);

                entity.Property(e => e.Totala)
                    .HasColumnName("totala")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Totalb)
                    .HasColumnName("totalb")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Vatbill)
                    .HasColumnName("vatbill")
                    .HasMaxLength(10);

                entity.Property(e => e.Wdper)
                    .HasColumnName("wdper")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Winsor).HasColumnName("winsor");

                entity.Property(e => e.Wonder)
                    .HasColumnName("wonder")
                    .HasMaxLength(10);

                entity.Property(e => e.WsRate)
                    .HasColumnName("ws_rate")
                    .HasMaxLength(40);

                entity.Property(e => e.Wsbk)
                    .HasColumnName("wsbk")
                    .HasMaxLength(50);

                entity.Property(e => e.Ydepdate)
                    .HasColumnName("ydepdate")
                    .HasMaxLength(10);

                entity.Property(e => e.Ydepperson)
                    .HasColumnName("ydepperson")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Ydeproute)
                    .HasColumnName("ydeproute")
                    .HasMaxLength(10);

                entity.Property(e => e.Yretdate)
                    .HasColumnName("yretdate")
                    .HasMaxLength(10);

                entity.Property(e => e.Yretperson)
                    .HasColumnName("yretperson")
                    .HasColumnType("decimal(18, 0)");

                entity.Property(e => e.Yretroute)
                    .HasColumnName("yretroute")
                    .HasMaxLength(10);
            });
        }
    }
}