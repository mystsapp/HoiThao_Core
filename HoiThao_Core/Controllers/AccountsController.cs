using Data;
using Data.Repository;
using Microsoft.AspNetCore.Mvc;
using System;

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
        public IActionResult Update(Account account)
        {
            if (!ModelState.IsValid)
            {
                return View(account);
            }

            var accountM = _accountRepository.GetById(account.Id);

            if (string.IsNullOrEmpty(account.Password))
                account.Password = accountM.Password;

            account.Ngaycapnhat = DateTime.Now;

            _accountRepository.Update(account);

            return RedirectToAction("List");
        }

        public IActionResult Delete(int id)
        {
            var account = _accountRepository.GetById(id);

            _accountRepository.Delete(account);

            return RedirectToAction("List");
        }
    }
}