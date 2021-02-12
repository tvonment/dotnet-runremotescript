using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;


namespace RunRemotePowershell.Models
{
    public class User
    {
        [Required(ErrorMessage = "Firstname is required")]
        public string FirstName { get; set; }

        [Required(ErrorMessage = "Lastname is required")]
        public string LastName { get; set; }

        [Required(ErrorMessage = "LogonName is required")]
        public string LogonName { get; set; }

        [Required(ErrorMessage = "DisplayName is required")]
        public string DisplayName { get; set; }

        [Required(ErrorMessage = "Email is required")]
        public string EMail { get; set; }
        public string PasswordResetNR { get; set; }
        public string TelephoneNR { get; set; }
        public string Description { get; set; }
        public string Company { get; set; }
        public string Title { get; set; }
        public string Kostenstelle { get; set; }
        public string Department { get; set; }
        public string Office { get; set; }

        [Required(ErrorMessage = "Manager is required")]
        public string Manager { get; set; }

        [Required(ErrorMessage = "Container is required")]
        public string Container { get; set; }
        public string Street { get; set; }
        public string City { get; set; }
        public string ZIP { get; set; }
        public string WebPage { get; set; }
        public string Gruppenmitgliedschaft { get; set; }
        public DateTime ExpirationDate { get; set; }
        public string License { get; set; }

        [Required(ErrorMessage = "ExchangeServer is required")]
        public string ExchangeServer { get; set; }

        [Required(ErrorMessage = "Password is required")]
        public string Password { get; set; }

    }
}
