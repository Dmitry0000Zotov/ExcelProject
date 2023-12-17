using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Domain.Entities
{
    public class Order
    {
        /// <summary>
        /// Код заявки
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Код товара
        /// </summary>
        public int ProductId { get; set; }

        /// <summary>
        /// Код клиента
        /// </summary>
        public int CustomerId { get; set; }

        /// <summary>
        /// Номер заявки
        /// </summary>
        public int OrderNum { get; set; }

        /// <summary>
        /// Требуемое количество
        /// </summary>
        public int Count { get; set; }

        /// <summary>
        /// Дата размещения
        /// </summary>
        public DateOnly DateCreate { get; set; }
    }
}
