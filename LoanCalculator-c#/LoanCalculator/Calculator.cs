using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LoanCalculator
{
    public class Calculator
    {
        public AnnualReport P, N, D, Now;
        public double InventoryTT, InventoryTD;//存货周转次数
        public double AdvanceCollectedTT, AdvanceCollectedTD;//预收账款周转次数
        public double AdvancePaymentTT, AdvancePaymentTD;//预付账款周转次数
        public double AccountPayableTT, AccountPayableTD;//应付账款周转次
        public double AccountReceivableTT, AccountReceivableTD;//应收账款周转次数
        public double AVE_Inventory;//平均存货余额
        public double AVE_AdvanceCollected;//平均预收账款余额
        public double AVE_AdvancePayment;//平均预付账款余额
        public double AVE_AccountPayable;//平均应付账款余额
        public double AVE_AccountReceivable;//平均应收账款余额
        public double ExpectedGrowthRate;
        public double WorkingCapital;//营运资金量
        public double WorkCapitalTurnover;//营运资金周转次数
        public double NewLoan;//新增流动资金贷款额度
        public double ExpectedIncome, NowMoneyStream, OtherMoneyStream, CurrentAssests, CurrentLiability;
        public double CustAmmount;//借款人自有资金;
        public void GetReport(AnnualReport P, AnnualReport N, AnnualReport D, AnnualReport Now)
        {
            this.P = P;
            this.N = N;
            this.D = D;
            this.Now = Now;
        }

        public double Gfp(double a)
        {
            return Convert.ToDouble(Math.Round((decimal)(a), 4, MidpointRounding.AwayFromZero));
        }
        public double Gtp(double a)
        {
            return Convert.ToDouble(Math.Round((decimal)(a), 2, MidpointRounding.AwayFromZero));
        }
        public double Gop(double a)
        {
            return Math.Ceiling(a);
        }

        //存货周次
        public void CalInventory()
        {
            AVE_Inventory = Math.Abs((P.Inventory + N.Inventory) / 2.0);
            InventoryTT = Gtp(D.CostRevenues / AVE_Inventory);
            InventoryTD = Gop(360 / InventoryTT);
        }
        public void CalInventory(double InventoryTT)
        {
            this.InventoryTT = InventoryTT;
            InventoryTD = Gop(360 / InventoryTT);
        }
        
        //预收账款周次
        public void CalAdvancCollected()
        {
            AVE_AdvanceCollected = Math.Abs((P.AdvanceCollected + N.AdvanceCollected) / 2.0);
            if (AVE_AdvanceCollected != 0)
            {
                AdvanceCollectedTT = Gtp(D.NetRevenues / AVE_AdvanceCollected);
                AdvanceCollectedTD = Gop(360 / AdvanceCollectedTT);
            }
            else
            {
                AdvanceCollectedTD = 0;
                AdvanceCollectedTT = 0;
            }
        }

        public void CalAdvancCollected(double AdvanceCollectedTT)
        {
            this.AdvanceCollectedTT = AdvanceCollectedTT;
            if (AVE_AdvanceCollected != 0)
            {
                AdvanceCollectedTT = Gtp(D.NetRevenues / AVE_AdvanceCollected);
                AdvanceCollectedTD = Gop(360 / AdvanceCollectedTT);
            }
            else
            {
                AdvanceCollectedTD = 0;
                AdvanceCollectedTT = 0;
            }
        }

        //预付账款周次
        public void CalAdvancePayment()
        {
            AVE_AdvancePayment = Math.Abs((P.AdvancePayment + N.AdvancePayment) / 2.0);
            if (AVE_AdvancePayment != 0)
            {
                AdvancePaymentTT = Gtp(D.CostRevenues / AVE_AdvancePayment);
                AdvancePaymentTD = Gop(360 / AdvancePaymentTT);
            }
            else
            {
                AdvancePaymentTD = 0;
                AdvancePaymentTT = 0;
            }
        }

        public void CalAdvancePayment(double AdvancePaymentTT)
        {
            this.AdvancePaymentTT = AdvancePaymentTT;
            if (AVE_AdvancePayment != 0)
            {
                AdvancePaymentTT = Gtp(D.CostRevenues / AVE_AdvancePayment);
                AdvancePaymentTD = Gop(360 / AdvancePaymentTT);
            }
            else
            {
                AdvancePaymentTD = 0;
                AdvancePaymentTT = 0;
            }
        }

        //应付账款周次
        public void CalAccountPayable()
        {
            AVE_AccountPayable = Math.Abs((P.AccountPayable + N.AccountPayable) / 2.0);
            AccountPayableTT = Gtp(D.CostRevenues / AVE_AccountPayable);
            AccountPayableTD = Gop(360 / AccountPayableTT);
        }

        public void CalAccountPayable(double AccountPayableTT)
        {
            this.AccountPayableTT = AccountPayableTT;
            AccountPayableTD = Gop(360 / AccountPayableTT);
        }

        //应收转款周次
        public void CalAccountReceivable()
        {
            AVE_AccountReceivable = Math.Abs((P.AccountReceivable + N.AccountReceivable) / 2.0);
            AccountReceivableTT = Gtp(D.NetRevenues / AVE_AccountReceivable);
            AccountReceivableTD = Gop(360 / AccountReceivableTT);
        }

        public void CalAccountReceivable(double AccountPayableTT)
        {
            this.AccountReceivableTT = AccountPayableTT;
            AccountReceivableTD = Gop(360 / AccountReceivableTT);
        }
        //营运资金周转次数
        public void CalWorkCapitalTurnover()
        {
            WorkCapitalTurnover = Gtp(360 / (InventoryTD + AccountReceivableTD - AccountPayableTD + AdvancePaymentTD - AdvanceCollectedTD));
        }

        //预计增长速率
        public void CalExpectedGrowthRate()
        {
            ExpectedGrowthRate = Gfp((ExpectedIncome - Now.NetRevenues) / Now.NetRevenues);
        }

        //营运资金量
        public void CalWorkCaptial()
        {
            WorkingCapital = Gop(Now.NetRevenues * (1 - Now.ProfitMargin) * (1 + ExpectedGrowthRate) / WorkCapitalTurnover);
        }

        //结果
        public void CalNewLoan()
        {
            NewLoan = WorkingCapital - CustAmmount - NowMoneyStream - OtherMoneyStream;
        }


        //首层计算
        public void FirstCal()
        {
            CalAccountPayable();
            CalAdvancCollected();
            CalAdvancePayment();
            CalAccountReceivable();
            CalInventory();

            CalWorkCapitalTurnover();
            CalExpectedGrowthRate();
            CalWorkCaptial();
            CalNewLoan();
        }

        //二层计算
        public void SecondCal()
        {
            CalWorkCapitalTurnover();
            CalExpectedGrowthRate();
            CalWorkCaptial();
            CalNewLoan();
        }
        //三层计算
        public void ThirdCal()
        {
            CalWorkCaptial();
            CalNewLoan();
        }

        //存货周次改变
        public void ChangeInventory(double Inv)
        {
            CalInventory(Inv);
            SecondCal();
        }

        //预付周次改变
        public void ChangeAdvancePayment(double adp)
        {
            CalAdvancePayment(adp);
            SecondCal();
        }

        //预收周次改变
        public void ChangeAdvanceCollected(double adc)
        {
            CalAdvancCollected(adc);
            SecondCal();
        }

        //应付周次改变
        public void ChangeAccountPayable(double apb)
        {
            CalAccountPayable(apb);
            SecondCal();
        }

        //应收周次改变
        public void ChangeAccountReceivable(double arb)
        {
            CalAccountReceivable(arb);
            SecondCal();
        }

        //营运周次改变
        public void ChangeWorkCapitalTurnOver(double wcto)
        {
            this.WorkCapitalTurnover = wcto;
            ThirdCal();
        }

        //利率改变
        public void ChangeProfitMargin(double pm)
        {
            Now.ProfitMargin = pm;
            ThirdCal();
        }
        
        //预计增长速率改变
        public void ChangeExpectedGrowthRate(double egr)
        {
            ExpectedGrowthRate = egr;
            ThirdCal();
        }
    }
}
