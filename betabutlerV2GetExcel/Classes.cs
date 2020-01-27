using System;
using System.Collections.Generic;

namespace betabutlerV2GetExcel
{
    public class Order
    {
        public DateTime date { get; set; }
        public string companyStatus { get; set; }
        public string name { get; set; }
        public string restaurant { get; set; }
        public string meal { get; set; }
        public double price { get; set; }
        public int quantaty { get; set; }
        public double grand { get; set; }
    }

    public class Day
    {
        public int daynumber { get; set; }
        public string name { get; set; }
        public List<Order> order { get; set; }
    }
    public class Person
    {
        public string name { get; set; }
        public List<Order> orders { get; set; }
    }
    public class Restaurant
    {
        public string restaurantName { get; set; }
        public List<Order> orders { get; set; }
    }
}
