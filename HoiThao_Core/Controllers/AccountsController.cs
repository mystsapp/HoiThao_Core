using Data;
using Data.Repository;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Linq;

namespace HoiThao_Core.Controllers
{
    public class AccountsController : Controller
    {
        private readonly IAccountRepository _accountRepository;

        public AccountsController(IAccountRepository accountRepository)
        {
            _accountRepository = accountRepository;
        }

        //[Route("Accounts")]
        public IActionResult List(string searchString, int page = 1)
        {
            ViewData["CurrentFilter"] = searchString;

            var accounts = _accountRepository.GetAccounts(searchString, page);

            ViewBag.Accounts = accounts;

            return View(accounts);
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Create(Account account)
        {
            if (!ModelState.IsValid)
            {
                return View(account);
            }

            account.Ngaytao = DateTime.Now;
            _accountRepository.Create(account);

            return RedirectToAction("List");
        }

        public IActionResult Update(int id)
        {
            var customer = _accountRepository.GetById(id);
            return View(customer);
        }

        [HttpPost]
        public IActionResult Update(Account account, string pass)
        {
            //var accountM = _accountRepository.GetById((int)(account.Id));
            //var oldPass = accountM.Password;

            if (string.IsNullOrEmpty(account.Password))
                account.Password = pass;

            account.Ngaycapnhat = DateTime.Now;

            if (ModelState.IsValid)
            {
                _accountRepository.Update(account);
                return RedirectToAction("List");

            }
            return View(account);
        }

        public IActionResult Delete(int id)
        {
            var account = _accountRepository.GetById(id);

            _accountRepository.Delete(account);

            return RedirectToAction("List");
        }
    }
}