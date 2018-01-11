#include <cstring>
using namespace std;
struct AnnualReport{
    string year;
    double ExpectedIncome;//预期收入
    double Inventory;//存货
    double CostRevenues;//营业成本
    double NetRevenues;//营业收入
    double AccountReceivable;//应收账款
    double AccountPayable;//应付账款
    double AdvancePayment;//预付账款
    double AdvanceCollected;//预收账款
    double NetProfit;//净利润
    double ProfitMargin;//利润率
    double CustAmmount;//借款人自有资金-货币资金
};

class Turnover {
private:
    string PRE_year, NOW_year, D_year;
    double InventoryTT ,InventoryTD;//存货周转次数
    double AdvanceCollectedTT, AdvanceCollectedTD;//预收账款周转次数
    double AdvancePaymentTT, AdvancePaymentTD;//预付账款周转次数
    double AccountPayableTT, AccountPayableTD;//应付账款周转次
    double AccountReceivableTT, AccountReceivableTD;//应收账款周转次数
    double CostRevenues;//营业成本
    double NetRevenues;//营业收入
    double AVE_Inventory;//平均存货余额
    double AVE_AdvanceCollected;//平均预收账款余额
    double AVE_AdvancePayment;//平均预付账款余额
    double AVE_AccountPayable;//平均应付账款余额
    double AVE_AccountReceivable;//平均应收账款余额
    double PRE_AccountReceivable, NOW_AccountReceivable;//应收账款
    double PRE_AccountPayable, NOW_AccountPayable;//应付账款
    double PRE_AdvancePayment, NOW_AdvancePayment;//预付账款
    double PRE_AdvanceCollected, NOW_AdvanceCollected;//预收账款
    double PRE_Inventory, NOW_Inventory;//存货
    double WorkingCapitalTurnover;
public:
    Turnover(AnnualReport, AnnualReport, AnnualReport);

    void setAVE_Inventory();
    void setAVE_AdvanceCollected();
    void setAVE_AdvancePayment();
    void setAVE_AccountPayable();
    void setAVE_AccountReceivable();

    void setInventoryTT();
    void setAdvanceCollectedTT();
    void setAdvancePaymentTT();
    void setAccountPayableTT();
    void setAccountReceivableTT();
    void setWorkCapitalTurnover();

    void Output(ofstream&);

    double getInventoryTD();
    double getAdvanceCollectedTD();
    double getAdvancePaymentTD();
    double getAccountPayableTD();
    double getAccountReceivableTD();
    double getWorkCapitalTurnover();
};

class CalOthers{
private:
    AnnualReport P, N, D;
    double ExpectedGrowthRate;//预计销售收入增长率
    double WorkingCapital;//营运资金量
    double WorkCapitalTurnover;//营运资金周转次数
    double NewLoan;//新增流动资金贷款额度
    double NowMoneyStream;
    double OtherMoneyStream;
public:
    CalOthers(AnnualReport, AnnualReport, AnnualReport, double, double, double);
    void setExpectedGrowthRate();
    void setWorkingCapital();
    void setNewLoan();
    double getWorkingCapital();
    double getExpectedGrowthRate();
    void Output(ofstream&);
};
