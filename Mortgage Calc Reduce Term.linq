<Query Kind="FSharpProgram">
  <NuGetReference>ExcelFinancialFunctions</NuGetReference>
  <Namespace>Excel.FinancialFunctions</Namespace>
</Query>

// Defaults, ignore
let fv = 0.00
let typ = PaymentDue.EndOfPeriod

// Mortgage details
let principal = 116250.00
let fixedRate = 2.95
let variableRate = 4.49
let term = 25.00
let fixedTerm = 5.00

// Pre-calculations
let numberOfPayments = term * 12.00
let numberOfFixedPayments = fixedTerm * 12.00
let numberOfVariablePayments = numberOfPayments - numberOfFixedPayments

let fixedRateMonthlyInterest = (fixedRate / 100.00) / 12.00
let variableRateMonthlyInterest = (variableRate / 100.00) / 12.00

// Overpayments
let overPayments = Map.ofList [
                       //(26.00, -500.00)
                       //(260.00, -2000.00)
                   ]

let fixedRatePeriodPayment = 
    (Financial.Pmt(fixedRateMonthlyInterest, numberOfPayments, principal, fv, typ))

// Schedule creation functions
let calculatePaymentValues interestRate numPeriods principal overPayment paymentAmount =
    let principal' = principal + overPayment
    let pmt = paymentAmount
    let ipmt = (principal' * interestRate) * -1.00
    let ppmt = paymentAmount - ipmt
    (pmt + overPayment, ipmt, ppmt + overPayment)

let monthlyInterestRate period =
    match period <= numberOfFixedPayments with
    | true -> fixedRateMonthlyInterest
    | false -> variableRateMonthlyInterest
    
let overPaymentAmount period = 
    match overPayments |> Map.tryFind period with
    | Some op -> op
    | None -> 0.00
    
let schedulePeriodData period = 
    let rate = monthlyInterestRate period
    let overPayment = overPaymentAmount period
    (period, ((numberOfPayments - period) + 1.00), rate, overPayment)

let createSchedulePeriod paymentAmount balance schedulePeriod =
    let (period, numPeriods, interestRate, overPayment) = schedulePeriod
    let (pmt, ipmt, ppmt) = calculatePaymentValues interestRate numPeriods balance overPayment paymentAmount
    let newBalance = balance + ppmt
    
    let (pmt', ppmt', newBalance') = match newBalance < 0.00 with
                                     | true -> let payment = balance * -1.00
                                               (payment, (payment + ipmt), (match newBalance > 0.00 with
                                                                            | true -> newBalance
                                                                            | false -> 0.00))
                                     | false -> (pmt, ppmt, newBalance)

    ((int period, pmt', ipmt, ppmt', newBalance'), newBalance)

let schedulePeriods = [1.00..numberOfPayments] |> List.map schedulePeriodData

let (schedule, bal) = schedulePeriods |> List.take 60 |>  List.mapFold (createSchedulePeriod fixedRatePeriodPayment) principal

let variableRatePeriodPayment = 
    (Financial.Pmt(variableRateMonthlyInterest, numberOfVariablePayments, bal, fv, typ))

let (schedule2, _) = schedulePeriods |> List.skip 60 |>  List.mapFold (createSchedulePeriod variableRatePeriodPayment) bal

// Display functions
let fmt (n:float) = n.ToString("0.00")

let writeScheduleRow schedulePeriod = 
    let (period, payment, interestPayment, principalPayment, balance) = schedulePeriod
    (period, (fmt payment), (fmt interestPayment), (fmt principalPayment), (fmt balance))
                    
schedule @ schedule2
|> List.filter (fun (_, pmt, _, _, _) -> pmt <= 0.00)
|> List.map writeScheduleRow
|> Dump
|> ignore