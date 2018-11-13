using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
namespace LoanCalculator
{
    public class AnnualReport
    {
        public string year;
        public double ExpectedIncome;//预期收入
        public double Inventory;//存货
        public double CostRevenues;//营业成本
        public double NetRevenues;//营业收入
        public double AccountReceivable;//应收账款
        public double AccountPayable;//应付账款
        public double AdvancePayment;//预付账款
        public double AdvanceCollected;//预收账款
        public double NetProfit;//净利润
        public double ProfitMargin;//利润率
        public double CurrentLiability;//流动负债
        public double CurrentAssests;//流动资金
      
    }
}
