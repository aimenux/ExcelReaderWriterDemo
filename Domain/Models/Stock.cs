using System.Collections.Generic;
using System.Linq;

namespace Domain.Models
{
    public class Stock
    {
        private readonly ICollection<Product> _products;

        public Stock(ICollection<Product> products)
        {
            _products = products;
        }

        public int TotalProducts => _products.Count;

        public decimal TotalPrice => _products.Sum(x => x.Price * x.Quantity);

        public override string ToString()
        {
            return $"{TotalProducts} Products costs {TotalPrice}$";
        }
    }
}
