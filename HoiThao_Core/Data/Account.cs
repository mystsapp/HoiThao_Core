﻿using System;
using System.ComponentModel.DataAnnotations;

namespace HoiThao_Core.Data
{
    public partial class Account
    {
        public int Id { get; set; }
        public string Nhom { get; set; }

        [Required]
        public string Username { get; set; }

        [Required]
        public string Password { get; set; }

        [Required]
        public string Hoten { get; set; }

        public string Daily { get; set; }
        public string Chinhanh { get; set; }
        public string Role { get; set; }
        public bool Doimatkhau { get; set; }
        public DateTime? Ngaydoimk { get; set; }
        public bool Trangthai { get; set; }
        public string Khoi { get; set; }
        public string Nguoitao { get; set; }

        [DataType(DataType.Date)]
        public DateTime? Ngaytao { get; set; }

        public string Nguoicapnhat { get; set; }
        public DateTime? Ngaycapnhat { get; set; }
    }
}