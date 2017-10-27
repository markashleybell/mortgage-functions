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
                       (26.00, -500.00)
                       // (31.00, -500.00)
                   ]

// Schedule creation functions
let calculatePaymentValues interestRate numPeriods principal overPayment period =
    let principal' = principal + overPayment
    let pmt = Financial.Pmt(interestRate, numPeriods, principal', fv, typ)
    let ipmt = Financial.IPmt(interestRate, period, numPeriods, principal', fv, typ)
    let ppmt = Financial.PPmt(interestRate, period, numPeriods, principal', fv, typ)
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

let createSchedulePeriod balance schedulePeriod =
    let (period, numPeriods, interestRate, overPayment) = schedulePeriod
    let (pmt, ipmt, ppmt) = calculatePaymentValues interestRate numPeriods balance overPayment 1.00
    let newBalance = balance + ppmt
    ((int period, pmt, ipmt, ppmt, newBalance), newBalance)

let schedulePeriods = [1.00..numberOfPayments] |> List.map schedulePeriodData

let (schedule, _) = schedulePeriods |> List.mapFold createSchedulePeriod principal

// Display functions
let fmt (n:float) = n.ToString("0.00")

let writeScheduleRow schedulePeriod = 
    let (period, payment, interestPayment, principalPayment, balance) = schedulePeriod
    (period, (fmt payment), (fmt interestPayment), (fmt principalPayment), (fmt balance))
                    
schedule
|> List.map writeScheduleRow
|> Dump
|> ignore
