using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GemboxExcel.Models
{
    public enum TransactionType
    {
        Debit,
        Credit
    }

    public class SampleTransaction
    {
        public TransactionType Type { get; set; }
        public string Description { get; set; }
    }
}
