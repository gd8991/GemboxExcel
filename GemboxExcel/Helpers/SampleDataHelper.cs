using GemboxExcel.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace GemboxExcel.Helpers
{
    public class SampleDataHelper
    {
        private static readonly List<SampleTransaction> _sampleTransactions;
        private static DataTable _sampleData = new DataTable();
        private static Random _amountRange = new Random(1000000);
        private static int _balance = 0;

        static SampleDataHelper()
        {
            _sampleTransactions = new List<SampleTransaction>();
            AddSampleTransactions();
        }

        public static DataTable GenerateSampleData(int noOfRows)
        {
            _sampleData = new DataTable();
            AddColumns();
            AddRows(noOfRows);
            return _sampleData;
        }

        private static void AddColumns()
        {
            _sampleData.Columns.Add("Seq. No.");
            _sampleData.Columns.Add("Transaction Type");
            _sampleData.Columns.Add("Description");
            _sampleData.Columns.Add("Amount");
            _sampleData.Columns.Add("Balance");
            _sampleData.Columns.Add("Remarks");
        }

        private static void AddRows(int noOfRows)
        {
            for (var idx = 1; idx <= noOfRows; idx++)
            {
                AddRow(idx);
            }
        }

        private static void AddRow(int idx)
        {
            var dataRow = _sampleData.NewRow();

            string transactionType;
            string description;
            int amount;

            if (idx == 1)
            {
                transactionType = "Debit";
                description = "Balance B/F";
                amount = _amountRange.Next(1, 1000000);
                _balance = amount;
            }
            else
            {
                var randomSampleTransaction = new Random();
                var transactionIdx = randomSampleTransaction.Next(_sampleTransactions.Count);
                var transaction = _sampleTransactions[transactionIdx];
                transactionType = transaction.Type.ToString();
                description = transaction.Description;
                var randomAmount = _amountRange.Next(1, 1000000);
                amount = randomAmount;
                var amount1 = transaction.Type == TransactionType.Credit ? amount : -1 * amount;
                _balance = _balance + amount1;
            }
            dataRow["Seq. No."] = idx.ToString();
            dataRow["Transaction Type"] = transactionType;
            dataRow["Description"] = description;
            dataRow["Amount"] = amount;
            dataRow["Balance"] = _balance;
            _sampleData.Rows.Add(dataRow);
        }

        private static void AddSampleTransactions()
        {
            _sampleTransactions.Add(new SampleTransaction { Type = TransactionType.Credit, Description = "Cash Deposited" });
            _sampleTransactions.Add(new SampleTransaction { Type = TransactionType.Credit, Description = "Cheque Deposited" });
            _sampleTransactions.Add(new SampleTransaction { Type = TransactionType.Credit, Description = "Reward Points" });
            _sampleTransactions.Add(new SampleTransaction { Type = TransactionType.Credit, Description = "Fuel Surcharge Waiver" });
            _sampleTransactions.Add(new SampleTransaction { Type = TransactionType.Debit, Description = "Cash Withdrawl for ATM" });
            _sampleTransactions.Add(new SampleTransaction { Type = TransactionType.Debit, Description = "Card Swipped" });
            _sampleTransactions.Add(new SampleTransaction { Type = TransactionType.Debit, Description = "ECS for EMI" });
            _sampleTransactions.Add(new SampleTransaction { Type = TransactionType.Debit, Description = "Cash Withdrawl" });
        }
    }
}
