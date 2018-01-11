#include <cstring>
#include <iostream>
#include <fstream>
#include "DataStructure.h"
#include <cmath>
using namespace std;
Turnover::Turnover(AnnualReport P, AnnualReport N, AnnualReport D) {
        PRE_year = P.year;
        NOW_year = N.year;
        D_year = D.year;
        PRE_AccountPayable = P.AccountPayable;
        PRE_AccountReceivable = P.AccountReceivable;
        PRE_AdvanceCollected = P.AdvanceCollected;
        PRE_AdvancePayment = P.AdvancePayment;
        PRE_Inventory = P.Inventory;
        NOW_AccountPayable = N.AccountPayable;
        NOW_AdvancePayment = N.AdvancePayment;
        NOW_AccountReceivable = N.AccountReceivable;
        NOW_AdvanceCollected = N.AdvanceCollected;
        NOW_Inventory = N.Inventory;
        CostRevenues = D.CostRevenues;
        NetRevenues = D.NetRevenues;

        setAVE_AdvanceCollected();
        setAVE_Inventory();
        setAVE_AccountPayable();
        setAVE_AccountReceivable();
        setAVE_AdvancePayment();

        setInventoryTT();
        setAccountReceivableTT();
        setAccountPayableTT();
        setAdvanceCollectedTT();
        setAdvancePaymentTT();
        setWorkCapitalTurnover();
}

void Turnover::setAVE_Inventory() {
        AVE_Inventory = abs((PRE_Inventory + NOW_Inventory) / 2.0);

}
void Turnover::setAVE_AdvanceCollected() {
        AVE_AdvanceCollected = abs((PRE_AdvanceCollected + NOW_AdvanceCollected) / 2.0);
}
void Turnover::setAVE_AdvancePayment() {
        AVE_AdvancePayment = abs((PRE_AdvancePayment + NOW_AdvancePayment) / 2.0);
}
void Turnover::setAVE_AccountPayable() {
        AVE_AccountPayable = abs((PRE_AccountPayable + NOW_AccountPayable) / 2.0);
}
void Turnover::setAVE_AccountReceivable() {
        AVE_AccountReceivable = abs((PRE_AccountReceivable + NOW_AccountReceivable) / 2.0);
}

void Turnover::setInventoryTT() {
        InventoryTT = CostRevenues / AVE_Inventory;
        InventoryTD = 360 / InventoryTT;
}//存货周转次数
void Turnover::setAdvanceCollectedTT() {
    if (AVE_AdvanceCollected != 0) {
        AdvanceCollectedTT = NetRevenues / AVE_AdvanceCollected;
        AdvanceCollectedTD = 360 / AdvanceCollectedTT;
    }  else {
        AdvanceCollectedTD = 0;
        AdvanceCollectedTT = 0;
    }

}//预收账款周转次数
void Turnover::setAdvancePaymentTT() {
    if (AVE_AdvancePayment != 0) {
        AdvancePaymentTT = CostRevenues / AVE_AdvancePayment;
        AdvancePaymentTD = 360 / AdvancePaymentTT;
    } else {
        AdvancePaymentTD = 0;
        AdvancePaymentTT = 0;
    }
}//预付账款周转次数

void Turnover::setAccountPayableTT() {
        AccountPayableTT = CostRevenues / AVE_AccountPayable;
        AccountPayableTD = 360 / AccountPayableTT;
}//应付账款周转次
void Turnover::setAccountReceivableTT() {
        AccountReceivableTT = NetRevenues / AVE_AccountReceivable;
        AccountReceivableTD = 360 / AccountReceivableTT;
}//应收账款周转次数
void Turnover::setWorkCapitalTurnover() {
        WorkingCapitalTurnover = 360 / (InventoryTD + AccountReceivableTD - AccountPayableTD + AdvancePaymentTD - AdvanceCollectedTD);
}

void Turnover::Output(ofstream& outFile) {
        outFile << "以" << PRE_year << "和" << NOW_year << "计算平均余额, 以"<<  D_year <<"为销售成本和营业收入 进行计算:"<< endl<<endl;
        outFile << "存货周转次数 = 营业成本 / 平均存货余额 = " << CostRevenues << " / " << AVE_Inventory << " = " << InventoryTT << ", "
                << "存货周转天数 = 360 / " << InventoryTT << " = " << InventoryTD << endl<< endl;
        outFile << "预收账款周转次数 = 营业收入 / 平均预收账款余额 = " << NetRevenues << " / " <<AVE_AdvanceCollected <<" = " << AdvanceCollectedTT << ", "
                << "预收账款周转天数 = 360 / " << AdvanceCollectedTT << " = " << AdvanceCollectedTD << endl<< endl;
        outFile << "预付账款周转次数 = 营业成本 / 平均预付账款余额 = " << CostRevenues << " / " <<AVE_AdvancePayment << " = " << AdvancePaymentTT << ", "
                << "预付账款周转天数 = 360 / " << AdvancePaymentTT << " = " << AdvancePaymentTD<< endl<< endl;
        outFile << "应付账款周转次数 = 销售成本 / 平均应付账款余额 = " << CostRevenues << " / " << AVE_AccountPayable << " = " << AccountPayableTT << ", "
                << "应付账款周转天数 = 360 / " << AccountPayableTT << " = " << AccountPayableTD<< endl<< endl;
        outFile << "应收账款周转次数 = 营业收入 / 平均应收账款余额 = " << NetRevenues << " / " << AVE_AccountReceivable << " = " << AccountReceivableTT << ", "
                << "应收账款周转天数 = 360 / " << AccountReceivableTT << " = " << AccountReceivableTD<< endl<< endl;
        outFile << "营运资金周转次数 = 360 / (存货周转天数 + 应收账款周转天数 -应付账款周转天数 + 预付账款周转天数 - 预收账款周转天数）= 360 / （" << InventoryTD
             << " + " << AccountReceivableTD <<" - "<< AccountPayableTD << " + " << AdvancePaymentTD <<" - " << AdvanceCollectedTD <<") = " << WorkingCapitalTurnover << endl<<endl;
}

double Turnover::getAccountPayableTD() {
    return AccountPayableTD;
}

double Turnover::getAccountReceivableTD() {
    return AccountReceivableTD;
}

double Turnover::getAdvanceCollectedTD() {
    return AdvanceCollectedTD;
}

double Turnover::getAdvancePaymentTD() {
    return AdvancePaymentTD;
}

double Turnover::getInventoryTD() {
    return InventoryTD;
}

double Turnover::getWorkCapitalTurnover() {
    return WorkingCapitalTurnover;
}






CalOthers::CalOthers(AnnualReport Pre, AnnualReport Now, AnnualReport Dev, double work, double Ns,double Os) {
    P = Pre;
    N = Now;
    D = Dev;
    WorkCapitalTurnover = work;
    this->NowMoneyStream = Ns;
    this->OtherMoneyStream = Os;
    setExpectedGrowthRate();
    setWorkingCapital();
    setNewLoan();
}

void CalOthers::setExpectedGrowthRate() {
    ExpectedGrowthRate = (D.ExpectedIncome - N.NetRevenues) / N.NetRevenues;
}

void CalOthers::setWorkingCapital() {
    WorkingCapital = N.NetRevenues * ( 1 - N.ProfitMargin) * (1 + ExpectedGrowthRate) / WorkCapitalTurnover;
}

double CalOthers::getExpectedGrowthRate() {
    return ExpectedGrowthRate;
}
double CalOthers::getWorkingCapital() {
    return WorkingCapital;
}
void CalOthers::setNewLoan() {
    NewLoan = WorkingCapital - D.CustAmmount - NowMoneyStream - OtherMoneyStream;
}

void CalOthers::Output(ofstream& outFile) {

    outFile << D.year << "年计划实现销售收入" << D.ExpectedIncome << "万元,预计销售收入年增长率" << "("
            <<D.ExpectedIncome << " - " <<N.NetRevenues << ")/" << N.NetRevenues << " = " << ExpectedGrowthRate * 100 << "%"<< endl<<endl;
    outFile << "营运资金量 = 上<" << N.year << ">年度销售收入 * ( 1 - 上<" << N.year <<">年度销售利润率)*(1+预计<" << D.year << ">销售收入年增长率/营运资金周转次数="<<
            N.NetRevenues << " * ( 1 - " << N.ProfitMargin * 100 << "% ) * ( 1 + " <<ExpectedGrowthRate * 100 << "% ) / " <<WorkCapitalTurnover << " = " << WorkingCapital << endl<< endl;
    outFile << "新增流动资金贷款额度 = 营运资金量 - 借款人自有资金 - 现有流动资金贷款 - 其他渠道提供的营运资金 = " << WorkingCapital << " - " << D.CustAmmount
            << " - " << NowMoneyStream  << " - " << OtherMoneyStream<< " = " << NewLoan << endl;
}