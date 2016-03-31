using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using Newtonsoft.Json.Serialization;
using Newtonsoft.Json;
namespace Microsoft_Graph_ASPNET_Excel_Donations.Models
{
    public class Donation
    {
        [JsonProperty("index")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime Date { get; set; }


        [Required]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:C}")]
        public decimal Amount { get; set; }

        [Required]
        public string Organization { get; set; }

        [Required]
        public string Month { get; set; }

        public Donation(
            DateTime date,
            decimal amount,
            string organization,
            string month)
        {
            Date = date;
            Amount = amount;
            Organization = organization;
            Month = month;
        }

        public Donation() { }
    }
}