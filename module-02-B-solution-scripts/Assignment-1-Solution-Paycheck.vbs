' Assignment Solution

option explicit

Const Name = "Adam"
Const AnnualSalary = 60000
Const RetirementContribution = 0.05
Const Taxrate = 0.20
Const MonthlyHealthPayment = 200

Dim GrossMonthlySalary, MonthlyTaxes, MonthlyRetirementContribution

Dim MonthlyNetSalary


GrossMonthlySalary = AnnualSalary / 12

MonthlyRetirementContribution = GrossMonthlySalary * RetirementContribution

MonthlyTaxes = GrossMonthlySalary * Taxrate

MonthlyNetSalary = GrossMonthlySalary - MonthlyRetirementContribution - MonthlyTaxes - MonthlyHealthPayment

MsgBox Name & " " & "receives $" & MonthlyNetSalary & " amount as a monthly salary after all deductions!", 64, "Result :"