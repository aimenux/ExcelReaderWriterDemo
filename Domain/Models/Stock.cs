using System;
using System.Collections.Generic;
using System.Linq;

namespace Domain.Models
{
    public class Stock
    {
        public IReadOnlyCollection<Product> Products { get; }

        public Stock(ICollection<Product> products)
        {
            Products = products?.ToList() ?? throw new ArgumentNullException(nameof(products));
        }

        public decimal TotalPrice => Products.Sum(x => x.Price * x.Quantity);

        public override string ToString()
        {
            return $"{Products.Count} Products costs {TotalPrice}$";
        }
    }
}
