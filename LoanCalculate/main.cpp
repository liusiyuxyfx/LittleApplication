#include <iostream>
#include <fstream>
#include "DataStructure.h"
#include <vector>
#include <sstream>
vector<vector<string>> strArray;
ofstream outFile;
double NowMoneyStream;
double OtherMoneyStream;
void OpenFile(){
    outFile.open("ans.txt", ios::out);

    ifstream inFile("data.csv", ios::in);
    string lineStr;
    while (getline(inFile, lineStr)){
        stringstream ss(lineStr);
        string str;
        vector<string> lineArray;
        // 按照逗号分隔
        while (getline(ss, str, ','))
            lineArray.push_back(str);
        strArray.push_back(lineArray);
    }

}

void ReadtoAP(AnnualReport *P, AnnualReport *N, AnnualReport *D) {
    P->year = strArray[0][1], N->year = strArray[0][2], D->year = strArray[0][3];
    P->CustAmmount = stod(strArray[1][1]), N->CustAmmount = stod(strArray[1][2]), D->CustAmmount = stod(strArray[1][3]);
    P->AccountPayable = stod(strArray[2][1]), N->AccountPayable = stod(strArray[2][2]), D->AccountPayable = stod(strArray[2][3]);
    P->AccountReceivable = stod(strArray[3][1]), N->AccountReceivable = stod(strArray[3][2]), D->AccountReceivable = stod(strArray[3][3]);
    P->AdvancePayment = stod(strArray[4][1]), N->AdvancePayment = stod(strArray[4][2]), D->AdvancePayment = stod(strArray[4][3]);
    P->AdvanceCollected = stod(strArray[5][1]), N->AdvanceCollected = stod(strArray[5][2]), D->AdvanceCollected = stod(strArray[5][3]);
    P->NetRevenues = stod(strArray[6][1]), N->NetRevenues = stod(strArray[6][2]), D->NetRevenues = stod(strArray[6][3]);
    P->CostRevenues = stod(strArray[7][1]), N->CostRevenues = stod(strArray[7][2]), D->CostRevenues = stod(strArray[7][3]);
    P->NetProfit = stod(strArray[8][1]), N->NetProfit = stod(strArray[8][2]), D->NetProfit = stod(strArray[8][3]);
    P->ProfitMargin = P->NetProfit / P->NetRevenues, N->ProfitMargin = N->NetProfit / N->NetRevenues, D->ProfitMargin = D->NetProfit / D->NetRevenues;
    P->Inventory = stod(strArray[9][1]), N->Inventory = stod(strArray[9][2]), D->Inventory = stod(strArray[9][3]);
    D->ExpectedIncome = stod(strArray[10][3]);
    NowMoneyStream = stod(strArray[11][1]);
    OtherMoneyStream = stod(strArray[12][1]);
}
int main() {
    OpenFile();
    AnnualReport P, N, D;
    ReadtoAP(&P, &N, &D);

    //P,N
    Turnover ansTPN(P, N, N);
    ansTPN.Output(outFile);
    CalOthers ansCPN(P, N, D, ansTPN.getWorkCapitalTurnover(), NowMoneyStream, OtherMoneyStream);
    ansCPN.Output(outFile);

    //N,D
    outFile << endl;
    Turnover ansTND(N, D, N);
    ansTND.Output(outFile);
    CalOthers ansCND(P, N, D, ansTND.getWorkCapitalTurnover(), NowMoneyStream, OtherMoneyStream);
    ansCND.Output(outFile);
}