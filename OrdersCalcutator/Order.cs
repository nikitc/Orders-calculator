using System;

namespace OrdersCalcutator
{
    public class Order
    {
        public string TradeName { get; set; }
        private string _tradeName {
            get
            {
                if (TradeName.StartsWith("#"))
                    return TradeName.Substring(1);
                if (TradeName.StartsWith("Заказ-"))
                    return TradeName.Substring(6);

                return TradeName;
            }
        }
        public string CompanyName { get; set; }
        public string MainContact { get; set; }
        public string CompanyContact { get; set; }
        public string Responsible { get; set; }
        public string TransactionStage { get; set; }
        public double Budget { get; set; }
        public DateTime CreateDate { get; set; }
        public string Creator { get; set; }
        public DateTime ChangeDate { get; set; }

        public string DeliveryAddress { get; set; }
        public DateTime? OrderDate { get; set; }

        public string WorkingEmail { get; set; }
        public string PrivateEmail { get; set; }

        public string WorkingPhone { get; set; }
        public string MobilePhone { get; set; }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            if (!(obj is Order m))
                return false;

            return m._tradeName == _tradeName;
        }

        public override int GetHashCode()
        {
            return _tradeName.GetHashCode();
        }
    }
}
