using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Domain.Entities
{
    public class Product
    {
        /// <summary>
        /// Код товара
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Наименование
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Единица измерения
        /// </summary>
        public string Uom { get; set; }

        /// <summary>
        /// Цена товара за единицу
        /// </summary>
        public decimal Price { get; set; }
    }
}
