<Query Kind="FSharpProgram">
  <NuGetReference>ExcelFinancialFunctions</NuGetReference>
  <Namespace>Excel.FinancialFunctions</Namespace>
</Query>

// So, what needs to happen is:
// Use the start values until a period where an overpayment is supposed to be made
// Subtract the overpayment value from the principal after all other calculations have been made
// Start again passing the new recalculated principal to calculatePayments (and the new number of payments)

// Defaults, ignore
let fv = 0.00
let typ = PaymentDue.EndOfPeriod

// Mortgage details
let principal = 116250.00
let fixedRate = 2.95
let variableRate = 4.49
let term = 25.00
let fixedTerm = 5.00

// Overpayments
let overPayments = Map.ofList [
                       (26.00, -500.00)
                       (31.00, -500.00)
                   ]

// Pre-calculations
let numberOfPayments = term * 12.00
let numberOfFixedPayments = fixedTerm * 12.00
let numberOfVariablePayments = numberOfPayments - numberOfFixedPayments

let fixedRateMonthlyInterest = (fixedRate / 100.00) / 12.00
let variableRateMonthlyInterest = (variableRate / 100.00) / 12.00

let calculatePayment rate numPeriods principal period =
    let pmt = Financial.Pmt(rate, numPeriods, principal, fv, typ)
    let ipmt = Financial.IPmt(rate, period, numPeriods, principal, fv, typ)
    let ppmt = Financial.PPmt(rate, period, numPeriods, principal, fv, typ)
    (pmt, ipmt, ppmt)

let periods = [1.00..numberOfPayments]
              |> List.map (fun i -> let rate = match i <= numberOfFixedPayments with
                                               | true -> fixedRateMonthlyInterest
                                               | false -> variableRateMonthlyInterest
                                    let overPayment = match overPayments |> Map.tryFind i with
                                                      | Some op -> op
                                                      | None -> 0.00
                                    (i, ((numberOfPayments - i) + 1.00), rate, overPayment))

// Display functions
let fmt (n:float) = n.ToString("0.00")

let writeRow row = 
    let (per, p, ip, pp, t) = row
    (per, (fmt p), (fmt ip), (fmt pp), (fmt t))

let (results, _) = (periods 
                    |> List.mapFold (fun balance paymentSlot ->
                                        let (period, numPeriods, interestRate, overPayment) = paymentSlot
                                        let (pmt, ipmt, ppmt) = calculatePayment interestRate numPeriods balance 1.00
                                        let newBalance = balance + ppmt + overPayment
                                        ((int period, pmt, ipmt, ppmt, newBalance), newBalance)) principal
                    )
                    
results
|> List.map writeRow
|> Dump
|> ignore



