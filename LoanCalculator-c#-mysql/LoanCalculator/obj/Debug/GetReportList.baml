   M S B A M L   `   `   `  ����I  ELoanCalculator, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null/clr-namespace:LoanCalculatorLoanCalculator  5   	   "    LoanCalculator.GetReportList   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   .localclr-namespace:LoanCalculator     5   	     � Title$  GetReportList��5   	   $	��450��6   $��300         ��6,   .��5
      �  � ColumnDefinitions 5   
   �� 5      $
��147* �6   �� 5      $	��24* �6   �� 5      $
��122* �6   .+�5   
   �� -   5         ��Name$ lv_ReportList��5      $	��250��6*   $	��300��66   $	��Top=�6C   $��Center=�6[   $��3��6x   $��22,0,21.6,0q�6�    �� View 5      �� 5      .P�5      �� $��id��6%    �� Width$ 50��61    �� DisplayMemberBinding 6<   ��  �� Path$ Id��� 5      $��名称��6%   $	 250��61    6=   �� $
 Name�5      5      �� 5   
   -   5         ��Name$ 	bt_Import��5      $��导入��6$   $
��Left=�61   $��
37,354,0,0q�6L   $	��Top=�6`   $��75��6x   6�   �� 5   
   -   5        $ 	bt_Delete��5      $��删除��6$   $
��Left=�61   $��23.8,354,0,0q�6L   $	��Top=�6b   $��75��6z   $��1��6�   $��2��6�   6�   5      5      