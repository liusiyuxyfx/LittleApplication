using System;
using System.Windows;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.Data;
namespace LoanCalculator
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    
    public partial class MainWindow : Window
    {
        public int Choose;
        public String ReportTitle;
        public AnnualReport P, N, D;
        double pntmp_InventoryTT;
        double pntmp_AccountPayableTT;
        double pntmp_AccountReceivableTT;
        double pntmp_AdvanceCollectedTT;
        double pntmp_AdvancePaymentTT;
        double pntmp_WorkCapitalTurnover;
        double pntmp_ProfitMargin;
        double pntmp_ExpectedGrowthRate;
        double ndtmp_InventoryTT;
        double ndtmp_AccountPayableTT;
        double ndtmp_AccountReceivableTT;
        double ndtmp_AdvanceCollectedTT;
        double ndtmp_AdvancePaymentTT;
        double ndtmp_WorkCapitalTurnover;
        double ndtmp_ProfitMargin;
        double ndtmp_ExpectedGrowthRate;
        public Calculator pnc, ndc;
        public double ExpectedIncome, NowMoneyStream, OtherMoneyStream, CurrentAssests, CurrentLiability;
        public double CustAmmount,Target;//借款人自有资金;     
        public int NowReportId;

        public void RefreshMethod()
        {
            //MessageBox.Show(NowReportId.ToString());
            importfrommysql();
        }

        public void pnChangeAv()
        {
            if (pnc.NewLoan > Target)
            {
                lbl_pnAccountPayableTT_av.Content = "下调";
                lbl_pnAccountReceivableTT_av.Content = "上调";
                lbl_pnAdvanceCollectedTT_av.Content = "下调";
                lbl_pnAdvancePaymenTT_av.Content = "上调";
                lbl_pnInventoryTT_av.Content = "上调";
                lbl_pnWorkingCapitalTurnoverTT_av.Content = "上调" ;
                lbl_pnProfitMargin_av.Content = "上调";
                lbl_pnExpectedGrowthRate_av.Content = "下调";

            } else
            {
                lbl_pnAccountPayableTT_av.Content = "上调";
                lbl_pnAccountReceivableTT_av.Content = "下调";
                lbl_pnAdvanceCollectedTT_av.Content = "上调";
                lbl_pnAdvancePaymenTT_av.Content = "下调";
                lbl_pnInventoryTT_av.Content = "下调";
                lbl_pnWorkingCapitalTurnoverTT_av.Content = "下调";
                lbl_pnProfitMargin_av.Content = "下调";
                lbl_pnExpectedGrowthRate_av.Content = "下调";
            }
        }
        public void ndChangeAv()
        {
            if (ndc.NewLoan > Target)
            {
                lbl_ndAccountPayableTT_av.Content = "下调";
                lbl_ndAccountReceivableTT_av.Content = "上调";
                lbl_ndAdvanceCollectedTT_av.Content = "下调";
                lbl_ndAdvancePaymenTT_av.Content = "上调";
                lbl_ndInventoryTT_av.Content = "上调";
                lbl_ndWorkingCapitalTurnoverTT_av.Content = "上调";
                lbl_ndProfitMargin_av.Content = "上调";
                lbl_ndExpectedGrowthRate_av.Content = "下调";

            }
            else
            {
                lbl_ndAccountPayableTT_av.Content = "上调";
                lbl_ndAccountReceivableTT_av.Content = "下调";
                lbl_ndAdvanceCollectedTT_av.Content = "上调";
                lbl_ndAdvancePaymenTT_av.Content = "下调";
                lbl_ndInventoryTT_av.Content = "下调";
                lbl_ndWorkingCapitalTurnoverTT_av.Content = "下调";
                lbl_ndProfitMargin_av.Content = "下调";
                lbl_ndExpectedGrowthRate_av.Content = "下调";
            }
        }
        public void pnFirstChange()
        {
            tb_pnWorkingCapitalTurnoverTT.Content = pnc.WorkCapitalTurnover.ToString();
            lbl_pnNewLoan.Content = "新增流动资金贷款额度:" + pnc.NewLoan.ToString() + "万元";
            pnChangeAv();
        }
        public void pnSecondChange()
        {
            lbl_pnNewLoan.Content = "新增流动资金贷款额度:" + pnc.NewLoan.ToString() + "万元";
            pnChangeAv();
        }
        public void ndFirstChange()
        {
            tb_ndWorkingCapitalTurnoverTT.Content = ndc.WorkCapitalTurnover.ToString();
            lbl_ndNewLoan.Content = "新增流动资金贷款额度:" + ndc.NewLoan.ToString() + "万元";
            ndChangeAv();
        }
        public void ndSecondChange()
        {
            lbl_ndNewLoan.Content = "新增流动资金贷款额度:" + ndc.NewLoan.ToString() + "万元";
            ndChangeAv();
        }

        //-----------------------------------------------------------------------------------------------------
        //预计增长率
        private void bt_pnExpectedGrowthRate_up(int x)
        {
            try
            {
                double tp = pnc.ExpectedGrowthRate;
                tp *= 100;
                tp += 0.01 * x;
                tb_pnExpectedGrowthRate.Content = tp.ToString();
                tp /= 100;
                pnc.ChangeExpectedGrowthRate(Gfp(tp));
                pnSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_pnExpectedGrowthRate_do(int x)
        {
            try
            {
                double tp = pnc.ExpectedGrowthRate;
                tp *= 100;
                tp -= 0.01 * x;
                tb_pnExpectedGrowthRate.Content = tp.ToString();
                tp /= 100;
                pnc.ChangeExpectedGrowthRate(Gfp(tp));
                pnSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_pnExpectedGrowthRate_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnExpectedGrowthRate_up(1);
        }

        private void bt_pnExpectedGrowthRate_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnExpectedGrowthRate_up(10);
        }

        private void bt_pnExpectedGrowthRate_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnExpectedGrowthRate_up(100);
        }

        private void bt_pnExpectedGrowthRate_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnExpectedGrowthRate_do(1);
        }

        private void bt_pnExpectedGrowthRate_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnExpectedGrowthRate_do(10);
        }

        private void bt_pnExpectedGrowthRate_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnExpectedGrowthRate_do(100);
        }

        private void bt_pnExpectedGrowthRate_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pntmp_ExpectedGrowthRate;
                tb_pnExpectedGrowthRate.Content = tp.ToString();
                pnc.ChangeExpectedGrowthRate(Gfp(tp / 100));
                pnSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        //-----------------------------------------------------------------------------------------------------
        //上年利率
        private void bt_pnProfitMargin_up(int x)
        {
            try
            {
                double tp = pnc.Now.ProfitMargin;
                tp *= 100;
                tp += 0.01 * x;
                tb_pnProfitMargin.Content = tp.ToString();
                tp /= 100;
                pnc.ChangeProfitMargin(Gfp(tp));
                pnSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_pnProfitMargin_do(int x)
        {
            try
            {
                double tp = pnc.Now.ProfitMargin;
                tp *= 100;
                tp -= 0.01 * x;
                tb_pnProfitMargin.Content = tp.ToString();
                tp /= 100;
                pnc.ChangeProfitMargin(Gfp(tp));
                pnSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_pnProfitMargin_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnProfitMargin_up(1);
        }

        private void bt_pnProfitMargin_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnProfitMargin_up(10);
        }

        private void bt_pnProfitMargin_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnProfitMargin_up(100);
        }

        private void bt_pnProfitMargin_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnProfitMargin_do(1);
        }

        private void bt_pnProfitMargin_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnProfitMargin_do(10);
        }

        private void bt_pnProfitMargin_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnProfitMargin_do(100);
        }

        private void bt_pnProfitMargin_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pntmp_ProfitMargin;
                tb_pnProfitMargin.Content = tp.ToString();
                pnc.ChangeProfitMargin(Gfp(tp / 100));
                pnSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        //-----------------------------------------------------------------------------------------------------
        //营运
        private void bt_pnWorkingCapitalTurnoverTT_up(int x)
        {
            try
            {
                double tp = pnc.WorkCapitalTurnover;
                tp += 0.01 * x;
                tb_pnWorkingCapitalTurnoverTT.Content = tp.ToString();
                pnc.ChangeWorkCapitalTurnOver(Gtp(tp));
                pnSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_pnWorkingCapitalTurnoverTT_do(int x)
        {
            try
            {
                double tp = pnc.WorkCapitalTurnover;
                tp -= 0.01 * x;
                tb_pnWorkingCapitalTurnoverTT.Content = tp.ToString();
                pnc.ChangeWorkCapitalTurnOver(Gtp(tp));
                pnSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }

        private void bt_pnWorkingCapitalTurnoverTT_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnWorkingCapitalTurnoverTT_up(1);
        }

        private void bt_pnWorkingCapitalTurnoverTT_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnWorkingCapitalTurnoverTT_up(10);
        }

        private void bt_pnWorkingCapitalTurnoverTT_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnWorkingCapitalTurnoverTT_up(100);
        }

        private void bt_pnWorkingCapitalTurnoverTT_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnWorkingCapitalTurnoverTT_do(1);
        }

        private void bt_pnWorkingCapitalTurnoverTT_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnWorkingCapitalTurnoverTT_do(10);
        }

        private void bt_pnWorkingCapitalTurnoverTT_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnWorkingCapitalTurnoverTT_do(100);
        }

        private void bt_pnWorkingCapitalTurnoverTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pntmp_WorkCapitalTurnover;
                tb_pnWorkingCapitalTurnoverTT.Content = tp.ToString();
                pnc.ChangeWorkCapitalTurnOver(Gtp(tp));
                pnSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }

        //-----------------------------------------------------------------------------------------------------
        //应收
        private void bt_pnAccountReceivableTT_up(int x)
        {
            try
            {
                double tp = pnc.AccountReceivableTT;
                tp += 0.01 * x;
                tb_pnAccountReceivableTT.Content = tp.ToString();
                pnc.ChangeAccountReceivable(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_pnAccountReceivableTT_do(int x)
        {
            try
            {
                double tp = pnc.AccountReceivableTT;
                tp -= 0.01 * x;
                tb_pnAccountReceivableTT.Content = tp.ToString();
                pnc.ChangeAccountReceivable(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_pnAccountReceivableTT_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountReceivableTT_up(1);
        }

        private void bt_pnAccountReceivableTT_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountReceivableTT_up(10);
        }

        private void bt_pnAccountReceivableTT_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountReceivableTT_up(100);
        }

        private void bt_pnAccountReceivableTT_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountReceivableTT_do(1);
        }

        private void bt_pnAccountReceivableTT_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountReceivableTT_do(10);
        }

        private void bt_pnAccountReceivableTT_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountReceivableTT_do(100);
        }

        private void bt_pnAccountReceivableTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pntmp_AccountReceivableTT;
                tb_pnAccountReceivableTT.Content = tp.ToString();
                pnc.ChangeAccountReceivable(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        //-----------------------------------------------------------------------------------------------------
        //应付
        private void bt_pnAccountPayableTT_up(int x)
        {
            try
            {
                double tp = pnc.AccountPayableTT;
                tp += 0.01 * x;
                tb_pnAccountPayableTT.Content = tp.ToString();
                pnc.ChangeAccountPayable(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_pnAccountPayableTT_do(int x)
        {
            try
            {
                double tp = pnc.AccountPayableTT;
                tp -= 0.01 * x;
                tb_pnAccountPayableTT.Content = tp.ToString();
                pnc.ChangeAccountPayable(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_pnAccountPayableTT_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountPayableTT_up(1);
        }

        private void bt_pnAccountPayableTT_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountPayableTT_up(10);
        }

        private void bt_pnAccountPayableTT_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountPayableTT_up(100);
        }

        private void bt_pnAccountPayableTT_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountPayableTT_do(1);
        }

        private void bt_pnAccountPayableTT_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountPayableTT_do(10);
        }

        private void bt_pnAccountPayableTT_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAccountPayableTT_do(100);
        }

        private void bt_pnAccountPayableTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pntmp_AccountPayableTT;
                tb_pnAccountPayableTT.Content = tp.ToString();
                pnc.ChangeAccountPayable(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        //-----------------------------------------------------------------------------------------------------
        //预付
        private void bt_pnAdvancePaymenTT_up(int x)
        {
            try
            {
                double tp = pnc.AdvancePaymentTT;
                if (tp == 0)
                {
                    MessageBox.Show("当预付和预收为0时，不能进行操作", "异常", MessageBoxButton.OK);
                    return;
                }
                tp += 0.01 * x;
                tb_pnAdvancePaymenTT.Content = tp.ToString();
                pnc.ChangeAdvancePayment(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_pnAdvancePaymenTT_do(int x)
        {
            try
            {
                double tp = pnc.AdvancePaymentTT;
                if (tp == 0)
                {
                    MessageBox.Show("当预付和预收为0时，不能进行操作", "异常", MessageBoxButton.OK);
                    return;
                }
                tp -= 0.01 * x;
                tb_pnAdvancePaymenTT.Content = tp.ToString();
                pnc.ChangeAdvancePayment(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_pnAdvancePaymenTT_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvancePaymenTT_up(1);
        }

        private void bt_pnAdvancePaymenTT_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvancePaymenTT_up(10);
        }

        private void bt_pnAdvancePaymenTT_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvancePaymenTT_up(100);
        }

        private void bt_pnAdvancePaymenTT_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvancePaymenTT_do(1);
        }

        private void bt_pnAdvancePaymenTT_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvancePaymenTT_do(10);
        }

        private void bt_pnAdvancePaymenTT_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvancePaymenTT_do(100);
        }

        private void bt_pnAdvancePaymenTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pntmp_AdvancePaymentTT;
                tb_pnAdvancePaymenTT.Content = tp.ToString();
                pnc.ChangeAdvancePayment(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }

        //-----------------------------------------------------------------------------------------------------
        //预收
        private void bt_pnAdvanceCoolectedTT_up(int x)
        {
            try
            {
                double tp = pnc.AdvanceCollectedTT;
                if (tp == 0)
                {
                    MessageBox.Show("当预付和预收为0时，不能进行操作", "异常", MessageBoxButton.OK);
                    return;
                }
                tp += 0.01 * x;
                tb_pnAdvanceCollectedTT.Content = tp.ToString();
                pnc.ChangeAdvanceCollected(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_pnAdvanceCoolectedTT_do(int x)
        {
            try
            {
                double tp = pnc.AdvanceCollectedTT;
                if (tp == 0)
                {
                    MessageBox.Show("当预付和预收为0时，不能进行操作", "异常", MessageBoxButton.OK);
                    return;
                }
                tp -= 0.01 * x;
                tb_pnAdvanceCollectedTT.Content = tp.ToString();
                pnc.ChangeAdvanceCollected(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_pnAdvanceCollectedTT_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvanceCoolectedTT_up(1);
        }

        private void bt_pnAdvanceCollectedTT_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvanceCoolectedTT_up(10);
        }

        private void bt_pnAdvanceCollectedTT_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvanceCoolectedTT_up(100);
        }

        private void bt_pnAdvanceCollectedTT_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvanceCoolectedTT_do(1);
        }

        private void bt_pnAdvanceCollectedTT_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvanceCoolectedTT_do(10);
        }

        private void bt_pnAdvanceCollectedTT_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_pnAdvanceCoolectedTT_do(100);
        }

        private void bt_pnAdvanceCollectedTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pntmp_AdvanceCollectedTT;
                tb_pnAdvanceCollectedTT.Content = tp.ToString();
                pnc.ChangeAdvanceCollected(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }


        //-------------------------------------------------------------------------------------------------------
        private void bt_pnInventoryTT_u0_Click(object sender, RoutedEventArgs e)
        {
            //pn鼠标点击事件
            //存货周次
            try
            {
                double tp = pnc.InventoryTT;
                tp += 0.01;
                tb_pnInventoryTT.Content = tp.ToString();
                pnc.ChangeInventory(Gtp(tp));
                pnFirstChange();
             } catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_pnInventoryTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pntmp_InventoryTT;
                tb_pnInventoryTT.Content = tp.ToString();
                pnc.ChangeInventory(tp);
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        private void bt_pnInventoryTT_u10_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pnc.InventoryTT;
                tp += 1;
                tb_pnInventoryTT.Content = tp.ToString();
                pnc.ChangeInventory(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        private void bt_pnInventoryTT_d10_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pnc.InventoryTT;
                tp -= 1;
                tb_pnInventoryTT.Content = tp.ToString();
                pnc.ChangeInventory(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        private void bt_pnInventoryTT_u1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pnc.InventoryTT;
                tp += 0.1;
                tb_pnInventoryTT.Content = tp.ToString();
                pnc.ChangeInventory(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_pnInventoryTT_d0_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pnc.InventoryTT;
                tp -= 0.01;
                tb_pnInventoryTT.Content = tp.ToString();
                pnc.ChangeInventory(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_pnInventoryTT_d1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = pnc.InventoryTT;
                tp -= 0.1;
                tb_pnInventoryTT.Content = tp.ToString();
                pnc.ChangeInventory(Gtp(tp));
                pnFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        //-----------------------------------------------------------------------------------------------------
        //预计增长率
        private void bt_ndExpectedGrowthRate_up(int x)
        {
            try
            {
                double tp = ndc.ExpectedGrowthRate;
                tp *= 100;
                tp += 0.01 * x;
                tb_ndExpectedGrowthRate.Content = tp.ToString();
                tp /= 100;
                ndc.ChangeExpectedGrowthRate(Gfp(tp));
                ndSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_ndExpectedGrowthRate_do(int x)
        {
            try
            {
                double tp = ndc.ExpectedGrowthRate;
                tp *= 100;
                tp -= 0.01 * x;
                tb_ndExpectedGrowthRate.Content = tp.ToString();
                tp /= 100;
                ndc.ChangeExpectedGrowthRate(Gfp(tp));
                ndSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_ndExpectedGrowthRate_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndExpectedGrowthRate_up(1);
        }

        private void bt_ndExpectedGrowthRate_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndExpectedGrowthRate_up(10);
        }

        private void bt_ndExpectedGrowthRate_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndExpectedGrowthRate_up(100);
        }

        private void bt_ndExpectedGrowthRate_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndExpectedGrowthRate_do(1);
        }

        private void bt_ndExpectedGrowthRate_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndExpectedGrowthRate_do(10);
        }

        private void bt_ndExpectedGrowthRate_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndExpectedGrowthRate_do(100);
        }

        private void bt_ndExpectedGrowthRate_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndtmp_ExpectedGrowthRate;
                tb_ndExpectedGrowthRate.Content = tp.ToString();
                ndc.ChangeExpectedGrowthRate(Gfp(tp / 100));
                ndSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        //-----------------------------------------------------------------------------------------------------
        //上年利率
        private void bt_ndProfitMargin_up(int x)
        {
            try
            {
                double tp = ndc.Now.ProfitMargin;
                tp *= 100;
                tp += 0.01 * x;
                tb_ndProfitMargin.Content = tp.ToString();
                tp /= 100;
                ndc.ChangeProfitMargin(Gfp(tp));
                ndSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_ndProfitMargin_do(int x)
        {
            try
            {
                double tp = ndc.Now.ProfitMargin;
                tp *= 100;
                tp -= 0.01 * x;
                tb_ndProfitMargin.Content = tp.ToString();
                tp /= 100;
                ndc.ChangeProfitMargin(Gfp(tp));
                ndSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_ndProfitMargin_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndProfitMargin_up(1);
        }

        private void bt_ndProfitMargin_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndProfitMargin_up(10);
        }

        private void bt_ndProfitMargin_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndProfitMargin_up(100);
        }

        private void bt_ndProfitMargin_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndProfitMargin_do(1);
        }

        private void bt_ndProfitMargin_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndProfitMargin_do(10);
        }

        private void bt_ndProfitMargin_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndProfitMargin_do(100);
        }

        private void bt_ndProfitMargin_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndtmp_ProfitMargin;
                tb_ndProfitMargin.Content = tp.ToString();
                ndc.ChangeProfitMargin(Gfp(tp / 100));
                ndSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        //-----------------------------------------------------------------------------------------------------
        //营运
        private void bt_ndWorkingCapitalTurnoverTT_up(int x)
        {
            try
            {
                double tp = ndc.WorkCapitalTurnover;
                tp += 0.01 * x;
                tb_ndWorkingCapitalTurnoverTT.Content = tp.ToString();
                ndc.ChangeWorkCapitalTurnOver(Gtp(tp));
                ndSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_ndWorkingCapitalTurnoverTT_do(int x)
        {
            try
            {
                double tp = ndc.WorkCapitalTurnover;
                tp -= 0.01 * x;
                tb_ndWorkingCapitalTurnoverTT.Content = tp.ToString();
                ndc.ChangeWorkCapitalTurnOver(Gtp(tp));
                ndSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }

        private void bt_ndWorkingCapitalTurnoverTT_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndWorkingCapitalTurnoverTT_up(1);
        }

        private void bt_ndWorkingCapitalTurnoverTT_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndWorkingCapitalTurnoverTT_up(10);
        }

        private void bt_ndWorkingCapitalTurnoverTT_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndWorkingCapitalTurnoverTT_up(100);
        }

        private void bt_ndWorkingCapitalTurnoverTT_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndWorkingCapitalTurnoverTT_do(1);
        }

        private void bt_ndWorkingCapitalTurnoverTT_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndWorkingCapitalTurnoverTT_do(10);
        }

        private void bt_ndWorkingCapitalTurnoverTT_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndWorkingCapitalTurnoverTT_do(100);
        }

        private void bt_ndWorkingCapitalTurnoverTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndtmp_WorkCapitalTurnover;
                tb_ndWorkingCapitalTurnoverTT.Content = tp.ToString();
                ndc.ChangeWorkCapitalTurnOver(Gtp(tp));
                ndSecondChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }

        //-----------------------------------------------------------------------------------------------------
        //应收
        private void bt_ndAccountReceivableTT_up(int x)
        {
            try
            {
                double tp = ndc.AccountReceivableTT;
                tp += 0.01 * x;
                tb_ndAccountReceivableTT.Content = tp.ToString();
                ndc.ChangeAccountReceivable(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_ndAccountReceivableTT_do(int x)
        {
            try
            {
                double tp = ndc.AccountReceivableTT;
                tp -= 0.01 * x;
                tb_ndAccountReceivableTT.Content = tp.ToString();
                ndc.ChangeAccountReceivable(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_ndAccountReceivableTT_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountReceivableTT_up(1);
        }

        private void bt_ndAccountReceivableTT_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountReceivableTT_up(10);
        }

        private void bt_ndAccountReceivableTT_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountReceivableTT_up(100);
        }

        private void bt_ndAccountReceivableTT_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountReceivableTT_do(1);
        }

        private void bt_ndAccountReceivableTT_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountReceivableTT_do(10);
        }

        private void bt_ndAccountReceivableTT_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountReceivableTT_do(100);
        }

        private void bt_ndAccountReceivableTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndtmp_AccountReceivableTT;
                tb_ndAccountReceivableTT.Content = tp.ToString();
                ndc.ChangeAccountReceivable(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        //-----------------------------------------------------------------------------------------------------
        //应付
        private void bt_ndAccountPayableTT_up(int x)
        {
            try
            {
                double tp = ndc.AccountPayableTT;
                tp += 0.01 * x;
                tb_ndAccountPayableTT.Content = tp.ToString();
                ndc.ChangeAccountPayable(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_ndAccountPayableTT_do(int x)
        {
            try
            {
                double tp = ndc.AccountPayableTT;
                tp -= 0.01 * x;
                tb_ndAccountPayableTT.Content = tp.ToString();
                ndc.ChangeAccountPayable(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_ndAccountPayableTT_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountPayableTT_up(1);
        }

        private void bt_ndAccountPayableTT_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountPayableTT_up(10);
        }

        private void bt_ndAccountPayableTT_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountPayableTT_up(100);
        }

        private void bt_ndAccountPayableTT_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountPayableTT_do(1);
        }

        private void bt_ndAccountPayableTT_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountPayableTT_do(10);
        }

        private void bt_ndAccountPayableTT_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAccountPayableTT_do(100);
        }

        private void bt_ndAccountPayableTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndtmp_AccountPayableTT;
                tb_ndAccountPayableTT.Content = tp.ToString();
                ndc.ChangeAccountPayable(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        //-----------------------------------------------------------------------------------------------------
        //预付
        private void bt_ndAdvancePaymenTT_up(int x)
        {
            try
            {
                double tp = ndc.AdvancePaymentTT;
                if (tp == 0)
                {
                    MessageBox.Show("当预付和预收为0时，不能进行操作", "异常", MessageBoxButton.OK);
                    return;
                }
                tp += 0.01 * x;
                tb_ndAdvancePaymenTT.Content = tp.ToString();
                ndc.ChangeAdvancePayment(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_ndAdvancePaymenTT_do(int x)
        {
            try
            {
                double tp = ndc.AdvancePaymentTT;
                if (tp == 0)
                {
                    MessageBox.Show("当预付和预收为0时，不能进行操作", "异常", MessageBoxButton.OK);
                    return;
                }
                tp -= 0.01 * x;
                tb_ndAdvancePaymenTT.Content = tp.ToString();
                ndc.ChangeAdvancePayment(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_ndAdvancePaymenTT_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvancePaymenTT_up(1);
        }

        private void bt_ndAdvancePaymenTT_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvancePaymenTT_up(10);
        }

        private void bt_ndAdvancePaymenTT_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvancePaymenTT_up(100);
        }

        private void bt_ndAdvancePaymenTT_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvancePaymenTT_do(1);
        }

        private void bt_ndAdvancePaymenTT_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvancePaymenTT_do(10);
        }

        private void bt_ndAdvancePaymenTT_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvancePaymenTT_do(100);
        }

        private void bt_ndAdvancePaymenTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndtmp_AdvancePaymentTT;
                tb_ndAdvancePaymenTT.Content = tp.ToString();
                ndc.ChangeAdvancePayment(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }

        //-----------------------------------------------------------------------------------------------------
        //预收
        private void bt_ndAdvanceCoolectedTT_up(int x)
        {
            try
            {
                double tp = ndc.AdvanceCollectedTT;
                if (tp == 0)
                {
                    MessageBox.Show("当预付和预收为0时，不能进行操作", "异常", MessageBoxButton.OK);
                    return;
                }
                tp += 0.01 * x;
                tb_ndAdvanceCollectedTT.Content = tp.ToString();
                ndc.ChangeAdvanceCollected(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_ndAdvanceCoolectedTT_do(int x)
        {
            try
            {
                double tp = ndc.AdvanceCollectedTT;
                if (tp == 0)
                {
                    MessageBox.Show("当预付和预收为0时，不能进行操作", "异常", MessageBoxButton.OK);
                    return;
                }
                tp -= 0.01 * x;
                tb_ndAdvanceCollectedTT.Content = tp.ToString();
                ndc.ChangeAdvanceCollected(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_ndAdvanceCollectedTT_u0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvanceCoolectedTT_up(1);
        }

        private void bt_ndAdvanceCollectedTT_u1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvanceCoolectedTT_up(10);
        }

        private void bt_ndAdvanceCollectedTT_u10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvanceCoolectedTT_up(100);
        }

        private void bt_ndAdvanceCollectedTT_d0_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvanceCoolectedTT_do(1);
        }

        private void bt_ndAdvanceCollectedTT_d1_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvanceCoolectedTT_do(10);
        }

        private void bt_ndAdvanceCollectedTT_d10_Click(object sender, RoutedEventArgs e)
        {
            bt_ndAdvanceCoolectedTT_do(100);
        }

        private void bt_ndAdvanceCollectedTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndtmp_AdvanceCollectedTT;
                tb_ndAdvanceCollectedTT.Content = tp.ToString();
                ndc.ChangeAdvanceCollected(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }


        //-------------------------------------------------------------------------------------------------------
        private void bt_ndInventoryTT_u0_Click(object sender, RoutedEventArgs e)
        {
            //nd鼠标点击事件
            //存货周次
            try
            {
                double tp = ndc.InventoryTT;
                tp += 0.01;
                tb_ndInventoryTT.Content = tp.ToString();
                ndc.ChangeInventory(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
        private void bt_ndInventoryTT_bc_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndtmp_InventoryTT;
                tb_ndInventoryTT.Content = tp.ToString();
                ndc.ChangeInventory(tp);
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        private void bt_ndInventoryTT_u10_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndc.InventoryTT;
                tp += 1;
                tb_ndInventoryTT.Content = tp.ToString();
                ndc.ChangeInventory(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        private void bt_ndInventoryTT_d10_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndc.InventoryTT;
                tp -= 1;
                tb_ndInventoryTT.Content = tp.ToString();
                ndc.ChangeInventory(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        private void bt_ndInventoryTT_u1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndc.InventoryTT;
                tp += 0.1;
                tb_ndInventoryTT.Content = tp.ToString();
                ndc.ChangeInventory(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_ndInventoryTT_d0_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndc.InventoryTT;
                tp -= 0.01;
                tb_ndInventoryTT.Content = tp.ToString();
                ndc.ChangeInventory(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }
        private void bt_ndInventoryTT_d1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double tp = ndc.InventoryTT;
                tp -= 0.1;
                tb_ndInventoryTT.Content = tp.ToString();
                ndc.ChangeInventory(Gtp(tp));
                ndFirstChange();
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        private void bt_GetndReport_Click(object sender, RoutedEventArgs e)
        {
            getFinalReport(ndc);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Choose = 1;
                MessageBox.Show("已选择" + txt_P_year.Text + "进行计算", "提醒", MessageBoxButton.OK);
            } catch(NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                Choose = 2;
                MessageBox.Show("已选择" + txt_N_year.Text + "进行计算", "提醒", MessageBoxButton.OK);
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                Choose = 3;
                MessageBox.Show("已选择" + txt_D_year.Text + "进行计算", "提醒", MessageBoxButton.OK);
            }
            catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }
        }

        private void bt_Save_Click(object sender, RoutedEventArgs e)
        {
            string connstring = "server=localhost;database=loancalculator;uid=root;pwd=root";
            MySqlConnection conn = new MySqlConnection(connstring);
            try
            {
                conn.Open();
                MessageBox.Show("连接成功", "测试结果");
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            MySqlCommand cmd = conn.CreateCommand();
            cmd.CommandText = "Insert into peryearreport" +
                 " (Cyear,AccountPayable,AccountReceivable,AdvancePayment,AdvanceCollected,NetRevenus,CostRevenus,NetProfit,Inventory) " +
                 " values(@Cyear,@AccountPayable,@AccountReceviable,@AdvancePayment,@AdvanceCollected,@NetRevenus,@CostRevenus,@NetProfit,@Inventory) ";
            MessageBox.Show(P.year.ToString());
            cmd.Parameters.AddWithValue("@Cyear", P.year);
            //MessageBox.Show(P.AccountPayable.ToString());
            cmd.Parameters.AddWithValue("@AccountPayable", P.AccountPayable);
            cmd.Parameters.AddWithValue("@AccountReceviable", P.AccountReceivable);
            cmd.Parameters.AddWithValue("@AdvancePayment",P.AdvancePayment);
            cmd.Parameters.AddWithValue("@AdvanceCollected",P.AdvanceCollected);
            cmd.Parameters.AddWithValue("@NetRevenus",P.NetRevenues);
            cmd.Parameters.AddWithValue("@CostRevenus",P.CostRevenues);
            cmd.Parameters.AddWithValue("@NetProfit",P.NetProfit);
            cmd.Parameters.AddWithValue("@Inventory",P.Inventory);
            cmd.ExecuteNonQuery();
            cmd.Cancel();
            long sonid1= cmd.LastInsertedId;

            cmd = conn.CreateCommand();
            cmd.CommandText = "Insert into peryearreport" +
                  " (Cyear,AccountPayable,AccountReceivable,AdvancePayment,AdvanceCollected,NetRevenus,CostRevenus,NetProfit,Inventory) " +
                  " values(@Cyear,@AccountPayable,@AccountReceviable,@AdvancePayment,@AdvanceCollected,@NetRevenus,@CostRevenus,@NetProfit,@Inventory) ";
            cmd.Parameters.AddWithValue("@Cyear", N.year);
            cmd.Parameters.AddWithValue("@AccountPayable", N.AccountPayable);
            cmd.Parameters.AddWithValue("@AccountReceviable", N.AccountReceivable);
            cmd.Parameters.AddWithValue("@AdvancePayment", N.AdvancePayment);
            cmd.Parameters.AddWithValue("@AdvanceCollected", N.AdvanceCollected);
            cmd.Parameters.AddWithValue("@NetRevenus", N.NetRevenues);
            cmd.Parameters.AddWithValue("@CostRevenus", N.CostRevenues);
            cmd.Parameters.AddWithValue("@NetProfit", N.NetProfit);
            cmd.Parameters.AddWithValue("@Inventory", N.Inventory);
            cmd.ExecuteNonQuery();
            cmd.Cancel();
            long sonid2 = cmd.LastInsertedId;

            cmd = conn.CreateCommand();
            cmd.CommandText = "Insert into peryearreport" +
                 " (Cyear,AccountPayable,AccountReceivable,AdvancePayment,AdvanceCollected,NetRevenus,CostRevenus,NetProfit,Inventory) " +
                 " values(@Cyear,@AccountPayable,@AccountReceviable,@AdvancePayment,@AdvanceCollected,@NetRevenus,@CostRevenus,@NetProfit,@Inventory) ";
            cmd.Parameters.AddWithValue("@Cyear", D.year);
            cmd.Parameters.AddWithValue("@AccountPayable", D.AccountPayable);
            cmd.Parameters.AddWithValue("@AccountReceviable", D.AccountReceivable);
            cmd.Parameters.AddWithValue("@AdvancePayment", D.AdvancePayment);
            cmd.Parameters.AddWithValue("@AdvanceCollected", D.AdvanceCollected);
            cmd.Parameters.AddWithValue("@NetRevenus", D.NetRevenues);
            cmd.Parameters.AddWithValue("@CostRevenus", D.CostRevenues);
            cmd.Parameters.AddWithValue("@NetProfit", D.NetProfit);
            cmd.Parameters.AddWithValue("@Inventory", D.Inventory);
            cmd.ExecuteNonQuery();
            cmd.Cancel();
            long sonid3 = cmd.LastInsertedId;

            cmd = conn.CreateCommand();
            cmd.CommandText = "Insert into annualreport" +
                "(ReportName,ExpectedIncome,NowMoneyStream,CurrentAssets,CurrentLiability,OtherMoneyStream,Target,Sonid1,sonid2,sonid3)" +
                " values(@ReportName,@ExpectedIncome,@NowMoneyStream,@CurrentAssets,@CurrentLiability,@OtherMoneyStream,@Target,@Sonid1,@sonid2,@sonid3)";
            cmd.Parameters.AddWithValue("@ReportName",txt_title.Text);
            cmd.Parameters.AddWithValue("@ExpectedIncome",this.ExpectedIncome);
            cmd.Parameters.AddWithValue("@NowMoneyStream",this.NowMoneyStream);
            cmd.Parameters.AddWithValue("@CurrentAssets",this.CurrentAssests);
            cmd.Parameters.AddWithValue("@CurrentLiability",this.CurrentLiability);
            cmd.Parameters.AddWithValue("@OtherMoneyStream",this.OtherMoneyStream);
            cmd.Parameters.AddWithValue("@Target",this.Target);
            cmd.Parameters.AddWithValue("@sonid1", sonid1);
            cmd.Parameters.AddWithValue("@sonid2", sonid2);
            cmd.Parameters.AddWithValue("@sonid3", sonid3);
            cmd.ExecuteNonQuery();
            cmd.Cancel();
            MessageBox.Show("插入成功");

        }

        public void importfrommysql()
        {
            string connstring = "server=localhost;database=loancalculator;uid=root;pwd=root";
            MySqlConnection conn = new MySqlConnection(connstring);
            try
            {
                conn.Open();
               // MessageBox.Show("连接成功", "测试结果");
            }
            catch (MySqlException ex)
            {
                //MessageBox.Show(ex.Message);
            }
            MySqlCommand cmd = conn.CreateCommand();
            cmd.CommandText = "select * from annualreport where reportid = @NowReportid";
            cmd.Parameters.AddWithValue("@NowReportid", NowReportId);
            cmd.ExecuteNonQuery();
            MySqlDataAdapter adap = new MySqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            adap.Fill(ds);
            cmd.Cancel();
            int[] sonid = new int[3];
            foreach (DataRow r in ds.Tables[0].Rows)
            {
                txt_title.Text = r["reportname"].ToString();
                txt_ExpectedIncome.Text = r["expectedincome"].ToString();
                txt_NowMoneyStream.Text = r["nowmoneystream"].ToString();
                txt_CurrentAssests.Text = r["currentassets"].ToString();
                txt_CurrentLiability.Text = r["currentliability"].ToString();
                txt_OtherMoneyStream.Text = r["othermoneystream"].ToString();
                txt_Target.Text = r["target"].ToString();
                sonid[0] = int.Parse(r["sonid1"].ToString());
                sonid[1] = int.Parse(r["sonid2"].ToString());
                sonid[2] = int.Parse(r["sonid3"].ToString());
            }
            cmd = conn.CreateCommand();
            cmd.CommandText = "select * from peryearreport where sonid = @sonid";
            cmd.Parameters.AddWithValue("@sonid",sonid[0]);
            cmd.ExecuteNonQuery();
            MySqlDataAdapter adap1 = new MySqlDataAdapter(cmd);
            DataSet ds1 = new DataSet();
            adap1.Fill(ds1);
            cmd.Cancel();
            foreach (DataRow r in ds1.Tables[0].Rows)
            {
                txt_P_year.Text = r["cyear"].ToString();
                txt_P_AccountPayable.Text = r["accountpayable"].ToString();
                txt_P_AccountReceivable.Text = r["accountreceivable"].ToString();
                txt_P_AdvancePayment.Text = r["advancePayment"].ToString();
                txt_P_AdvanceCollected.Text = r["advancecollected"].ToString();
                txt_P_NetRevenues.Text = r["netrevenus"].ToString();
                txt_P_CostRevenues.Text = r["costrevenus"].ToString();
                txt_P_NetProfit.Text = r["netprofit"].ToString();
                txt_P_Inventory.Text = r["inventory"].ToString();

            }
            cmd = conn.CreateCommand();
            cmd.CommandText = "select * from peryearreport where sonid = @sonid";
            cmd.Parameters.AddWithValue("@sonid", sonid[1]);
            cmd.ExecuteNonQuery();
            MySqlDataAdapter adap2 = new MySqlDataAdapter(cmd);
            DataSet ds2 = new DataSet();
            adap2.Fill(ds2);
            cmd.Cancel();
            foreach (DataRow r in ds2.Tables[0].Rows)
            {
                txt_N_year.Text = r["cyear"].ToString();
                txt_N_AccountPayable.Text = r["accountpayable"].ToString();
                txt_N_AccountReceivable.Text = r["accountreceivable"].ToString();
                txt_N_AdvancePayment.Text = r["advancePayment"].ToString();
                txt_N_AdvanceCollected.Text = r["advancecollected"].ToString();
                txt_N_NetRevenues.Text = r["netrevenus"].ToString();
                txt_N_CostRevenues.Text = r["costrevenus"].ToString();
                txt_N_NetProfit.Text = r["netprofit"].ToString();
                txt_N_Inventory.Text = r["inventory"].ToString();

            }
            cmd = conn.CreateCommand();
            cmd.CommandText = "select * from peryearreport where sonid = @sonid";
            cmd.Parameters.AddWithValue("@sonid", sonid[2]);
            cmd.ExecuteNonQuery();
            MySqlDataAdapter adap3 = new MySqlDataAdapter(cmd);
            DataSet ds3 = new DataSet();
            adap3.Fill(ds3);
            cmd.Cancel();
            foreach (DataRow r in ds3.Tables[0].Rows)
            {
                txt_D_year.Text = r["cyear"].ToString();
                txt_D_AccountPayable.Text = r["accountpayable"].ToString();
                txt_D_AccountReceivable.Text = r["accountreceivable"].ToString();
                txt_D_AdvancePayment.Text = r["advancePayment"].ToString();
                txt_D_AdvanceCollected.Text = r["advancecollected"].ToString();
                txt_D_NetRevenues.Text = r["netrevenus"].ToString();
                txt_D_CostRevenues.Text = r["costrevenus"].ToString();
                txt_D_NetProfit.Text = r["netprofit"].ToString();
                txt_D_Inventory.Text = r["inventory"].ToString();

            }
            getReportValues();
        }

        private void bt_import_Click(object sender, RoutedEventArgs e)
        {
            GetReportList window1 = new GetReportList();
            window1.PassValuesEvent += new GetReportList.PassValuesHandler(ReceiveValues);
            window1.Owner = this;
            window1.ShowDialog();

        }

        private void bt_GetpnReport_Click(object sender, RoutedEventArgs e)
        {
            getFinalReport(pnc);
        }

        private void ReceiveValues(object sender, PassValuesEventArgs e)
        {
            this.NowReportId = e.Value;
            MessageBox.Show(e.Value.ToString());
        }

        //-------------------------------------------------------------------------------------------------------
        //生成报告
        public void getFinalReport(Calculator cl)
        {
            try
            {
                txt_FinalReport.Clear();
                txt_FinalReport.Text += "存货周转次数 = 营业成本 / 平均存货余额 = " + cl.D.CostRevenues.ToString() + " / "
                    + cl.AVE_Inventory.ToString() + " = " + cl.InventoryTT.ToString() + ", 存货周转天数 = 360 / "
                    + cl.InventoryTT.ToString() + " = " + cl.InventoryTD.ToString() + Environment.NewLine;
                txt_FinalReport.Text += "预收账款周转次数 = 营业收入 / 平均预收账款余额 = " + cl.D.NetRevenues.ToString() + " / "
                    + cl.AVE_AdvanceCollected.ToString() + " = " + cl.AdvanceCollectedTT.ToString() + ", 预收账款周转天数 = 360 / "
                    + cl.AdvanceCollectedTT.ToString() + " = " + cl.AdvanceCollectedTD.ToString() + Environment.NewLine;
                txt_FinalReport.Text += "预付账款周转次数 = 营业成本 / 平均预付账款余额 = " + cl.D.CostRevenues.ToString() + " / "
                    + cl.AVE_AdvancePayment.ToString() + " = " + cl.AdvancePaymentTT.ToString() + ", 预付账款周转天数 = 360 / "
                    + cl.AdvancePaymentTT.ToString() + " = " + cl.AdvancePaymentTD.ToString() + Environment.NewLine;
                txt_FinalReport.Text += "应付账款周转次数 = 销售成本 / 平均应付账款余额 = " + cl.D.CostRevenues.ToString() + " / "
                    + cl.AVE_AccountPayable.ToString() + " = " + cl.AccountPayableTT.ToString() + ", 应付账款周转天数 = 360 / "
                    + cl.AccountPayableTT.ToString() + " = " + cl.AccountPayableTD.ToString() + Environment.NewLine;
                txt_FinalReport.Text += "应收账款周转次数 = 营业收入 / 平均应收账款余额 = " + cl.D.NetRevenues.ToString() + " / "
                    + cl.AVE_AccountReceivable.ToString() + " = " + cl.AccountReceivableTT.ToString() + ", 应收账款周转天数 = 360 / "
                    + cl.AccountReceivableTT.ToString() + " = " + cl.AccountReceivableTD.ToString() + Environment.NewLine;
                txt_FinalReport.Text += "营运资金周转次数 = 360 / (存货周转天数 + 应收账款周转天数 -应付账款周转天数 + 预付账款周转天数 - 预收账款周转天数）= 360 /（"
                    + cl.InventoryTD.ToString() + " + " + cl.AccountReceivableTD.ToString() + " - " + cl.AccountPayableTD.ToString()
                    + " + " + cl.AdvancePaymentTD.ToString() + " - " + cl.AdvanceCollectedTD + " ) = " + cl.WorkCapitalTurnover.ToString() + Environment.NewLine;
                txt_FinalReport.Text += "营运资金量 = 上年度销售收入 × ( 1 - 上年度销售利润率 ) × ( 1 + 预计销售收入年增长率 ） / 营运资金周转次数 = "
                    + cl.Now.NetRevenues.ToString() + " × ( 1 - " + cl.Now.ProfitMargin.ToString() + " ) × ( 1 + " + cl.ExpectedGrowthRate.ToString()
                    + " ) / " + cl.WorkCapitalTurnover.ToString() + " = " + cl.WorkingCapital.ToString() + Environment.NewLine;
                txt_FinalReport.Text += "借款人自有资金 = 流动资金 - 流动负债 = " + cl.CurrentAssests.ToString() + " - " + cl.CurrentLiability.ToString()
                    + " = " + cl.CustAmmount.ToString() + Environment.NewLine;
                txt_FinalReport.Text += "新增流动资金贷款额度 = 营运资金量 - 借款人自有资金 - 现有流动资金贷款 - 其他渠道提供的营运资金 ="
                    + cl.WorkingCapital.ToString() + " - " + cl.CustAmmount.ToString() + " - " + cl.NowMoneyStream.ToString() +
                    " - " + cl.OtherMoneyStream.ToString() + " = " + cl.NewLoan.ToString() + Environment.NewLine;
            } catch (NullReferenceException e2)
            {
                MessageBox.Show("请先输入报告", "异常", MessageBoxButton.OK);
            }

        }
    
        public MainWindow()
        {
            InitializeComponent();
            P = new AnnualReport();
            N = new AnnualReport();
            D = new AnnualReport();
            pnc = new Calculator();
            ndc = new Calculator();
            Choose = 2;
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
            return Convert.ToDouble(Math.Round((decimal)(a), 0, MidpointRounding.AwayFromZero));
        }

        public void test()
        {
            txt_P_year.Text = "2013";
            txt_N_year.Text = "2014";
            txt_D_year.Text = "2015";

            txt_P_AccountPayable.Text = "2743.1";
            txt_N_AccountPayable.Text = "3074";
            txt_D_AccountPayable.Text = "881";

            txt_P_AccountReceivable.Text = "15758";
            txt_N_AccountReceivable.Text = "12598";
            txt_D_AccountReceivable.Text = "10532";

            txt_P_AdvancePayment.Text = txt_N_AdvancePayment.Text = txt_D_AdvancePayment.Text = "0";
            txt_P_AdvanceCollected.Text = txt_N_AdvanceCollected.Text = txt_D_AdvanceCollected.Text = "0";

            txt_P_NetRevenues.Text = "30328";
            txt_N_NetRevenues.Text = "33821";
            txt_D_NetRevenues.Text = "12069";

            txt_P_CostRevenues.Text = "25395";
            txt_N_CostRevenues.Text = "28583";
            txt_D_CostRevenues.Text = "10084";


            txt_P_NetProfit.Text = "938";
            txt_N_NetProfit.Text = "1250";
            txt_D_NetProfit.Text = "347";

            txt_P_Inventory.Text = "1916";
            txt_N_Inventory.Text = "2045";
            txt_D_Inventory.Text = "4017";

            txt_ExpectedIncome.Text = "34500";
            txt_NowMoneyStream.Text = "10873";
            txt_OtherMoneyStream.Text = "0";
            txt_CurrentAssests.Text = "1000";
            txt_CurrentLiability.Text = "400";
            txt_Target.Text = "300";
            Connector conn = new Connector();
            conn.connect();
        }
        
        private void getReportValues()
        {
            this.ReportTitle = txt_title.Text;
            P.year = txt_P_year.Text;
            N.year = txt_N_year.Text;
            D.year = txt_D_year.Text;

            P.AccountPayable = double.Parse(txt_P_AccountPayable.Text);
            N.AccountPayable = double.Parse(txt_N_AccountPayable.Text);
            D.AccountPayable = double.Parse(txt_D_AccountPayable.Text);

            P.AccountReceivable = double.Parse(txt_P_AccountReceivable.Text);
            N.AccountReceivable = double.Parse(txt_N_AccountReceivable.Text);
            D.AccountReceivable = double.Parse(txt_D_AccountReceivable.Text);

            P.AdvancePayment = double.Parse(txt_P_AdvancePayment.Text);
            N.AdvancePayment = double.Parse(txt_N_AdvancePayment.Text);
            D.AdvancePayment = double.Parse(txt_D_AdvancePayment.Text);

            P.AdvanceCollected = double.Parse(txt_P_AdvanceCollected.Text);
            N.AdvanceCollected = double.Parse(txt_N_AdvanceCollected.Text);
            D.AdvanceCollected = double.Parse(txt_D_AdvanceCollected.Text);

            P.NetRevenues = double.Parse(txt_P_NetRevenues.Text);
            N.NetRevenues = double.Parse(txt_N_NetRevenues.Text);
            D.NetRevenues = double.Parse(txt_D_NetRevenues.Text);

            P.CostRevenues = double.Parse(txt_P_CostRevenues.Text);
            N.CostRevenues = double.Parse(txt_N_CostRevenues.Text);
            D.CostRevenues = double.Parse(txt_D_CostRevenues.Text);

            P.NetProfit = double.Parse(txt_P_NetProfit.Text);
            N.NetProfit = double.Parse(txt_N_NetProfit.Text);
            D.NetProfit = double.Parse(txt_D_NetProfit.Text);

            P.Inventory = double.Parse(txt_P_Inventory.Text);
            N.Inventory = double.Parse(txt_N_Inventory.Text);
            D.Inventory = double.Parse(txt_D_Inventory.Text);

            P.ProfitMargin = Gfp(P.NetProfit / P.NetRevenues);
            N.ProfitMargin = Gfp(N.NetProfit / N.NetRevenues);
            D.ProfitMargin = Gfp(D.NetProfit / D.NetRevenues);

            Target = double.Parse(txt_Target.Text);

            this.ExpectedIncome = double.Parse(txt_ExpectedIncome.Text);
            this.NowMoneyStream = double.Parse(txt_NowMoneyStream.Text);
            this.OtherMoneyStream = double.Parse(txt_OtherMoneyStream.Text);
            this.CurrentAssests = double.Parse(txt_CurrentAssests.Text);
            this.CurrentLiability = double.Parse(txt_CurrentLiability.Text);
            this.CustAmmount = this.CurrentAssests - this.CurrentLiability;
        }
        private void bt_GetReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //test();
                getReportValues();



                if (Choose == 1)
                {
                    pnc.GetReport(P, N, P, P);
                    ndc.GetReport(N, D, P, P);
                } else if (Choose == 2)
                {
                    pnc.GetReport(P, N, N, N);
                    ndc.GetReport(N, D, N, N);
                } else if (Choose == 3)
                {
                    pnc.GetReport(P, N, D, D);
                    ndc.GetReport(N, D, D, D);
                }
                
                pnc.ExpectedIncome = ndc.ExpectedIncome = this.ExpectedIncome;
                pnc.NowMoneyStream = ndc.NowMoneyStream = this.NowMoneyStream;
                pnc.OtherMoneyStream = ndc.OtherMoneyStream = this.OtherMoneyStream;
                pnc.CurrentAssests = ndc.CurrentAssests = this.CurrentAssests;
                pnc.CurrentLiability = ndc.CurrentLiability = this.CurrentLiability;
                pnc.CustAmmount = ndc.CustAmmount = this.CustAmmount;

                pnc.FirstCal();
                ndc.FirstCal();

                pntmp_InventoryTT = pnc.InventoryTT;
                pntmp_AccountPayableTT = pnc.AccountPayableTT;
                pntmp_AccountReceivableTT = pnc.AccountReceivableTT;
                pntmp_AdvanceCollectedTT = pnc.AdvanceCollectedTT;
                pntmp_AdvancePaymentTT = pnc.AdvancePaymentTT;
                pntmp_WorkCapitalTurnover = pnc.WorkCapitalTurnover;
                pntmp_ProfitMargin = pnc.Now.ProfitMargin * 100;
                pntmp_ExpectedGrowthRate = pnc.ExpectedGrowthRate * 100;
                ndtmp_InventoryTT = ndc.InventoryTT;
                ndtmp_AccountPayableTT = ndc.AccountPayableTT;
                ndtmp_AccountReceivableTT = ndc.AccountReceivableTT;
                ndtmp_AdvanceCollectedTT = ndc.AdvanceCollectedTT;
                ndtmp_AdvancePaymentTT = ndc.AdvancePaymentTT;
                ndtmp_WorkCapitalTurnover = ndc.WorkCapitalTurnover;
                ndtmp_ProfitMargin = ndc.Now.ProfitMargin * 100;
                ndtmp_ExpectedGrowthRate = ndc.ExpectedGrowthRate * 100;

                lbl_pnTitle.Content = "以" + P.year + "年与" + N.year + "年为基础计算平均值";
                lbl_pnInventoryTT_or.Content = tb_pnInventoryTT.Content = pnc.InventoryTT.ToString();
                lbl_pnAccountPayableTT_or.Content = tb_pnAccountPayableTT.Content = pnc.AccountPayableTT.ToString();
                lbl_pnAccountReceivableTT_or.Content = tb_pnAccountReceivableTT.Content = pnc.AccountReceivableTT.ToString();
                lbl_pnAdvanceCollectedTT_or.Content = tb_pnAdvanceCollectedTT.Content = pnc.AdvanceCollectedTT.ToString();
                lbl_pnAdvancePaymenTT_or.Content = tb_pnAdvancePaymenTT.Content = pnc.AdvancePaymentTT.ToString();
                lbl_pnWorkingCapitalTurnoverTT_or.Content = tb_pnWorkingCapitalTurnoverTT.Content = pnc.WorkCapitalTurnover.ToString();
                lbl_pnProfitMargin_or.Content = tb_pnProfitMargin.Content = pntmp_ProfitMargin.ToString();
                lbl_pnExpectedGrowthRate_or.Content = tb_pnExpectedGrowthRate.Content = pntmp_ExpectedGrowthRate.ToString();
                lbl_pnNewLoan.Content = "新增流动资金贷款额度:" + pnc.NewLoan.ToString() + "万元";

                lbl_ndTitle.Content = "以" + N.year + "年与" + D.year + "年为基础计算平均值";
                lbl_ndInventoryTT_or.Content = tb_ndInventoryTT.Content = ndc.InventoryTT.ToString();
                lbl_ndAccountPayableTT_or.Content = tb_ndAccountPayableTT.Content = ndc.AccountPayableTT.ToString();
                lbl_ndAccountReceivableTT_or.Content = tb_ndAccountReceivableTT.Content = ndc.AccountReceivableTT.ToString();
                lbl_ndAdvanceCollectedTT_or.Content = tb_ndAdvanceCollectedTT.Content = ndc.AdvanceCollectedTT.ToString();
                lbl_ndAdvancePaymenTT_or.Content = tb_ndAdvancePaymenTT.Content = ndc.AdvancePaymentTT.ToString();
                lbl_ndWorkingCapitalTurnoverTT_or.Content = tb_ndWorkingCapitalTurnoverTT.Content = ndc.WorkCapitalTurnover.ToString();
                lbl_ndProfitMargin_or.Content = tb_ndProfitMargin.Content = ndc.Now.ProfitMargin.ToString();
                lbl_ndExpectedGrowthRate_or.Content = tb_ndExpectedGrowthRate.Content = ndtmp_ExpectedGrowthRate.ToString();
                lbl_ndNewLoan.Content = "新增流动资金贷款额度:" + ndc.NewLoan.ToString() + "万元";

            }
            catch (FormatException e1)
            {
                MessageBox.Show("有某些值未输入", "异常", MessageBoxButton.OK);
            }
            
        }


            
    }
}
