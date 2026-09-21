Public Class BankAccount

    ' Private field that stores the account balance
    Private _balance As Decimal

    ' Read-only property that allows the balance to be viewed
    ' but not changed directly
    Public ReadOnly Property Balance As Decimal
        Get
            Return _balance
        End Get
    End Property

    ' Constructor to set the initial balance
    Public Sub New(Optional initialBalance As Decimal = 0D)
        If initialBalance < 0D Then
            Throw New ArgumentException("Initial balance cannot be negative.")
        End If

        _balance = initialBalance
    End Sub

    ' Deposits money into the account
    Public Sub Deposit(amount As Decimal)
        If amount < 0D Then
            Throw New ArgumentException("Deposit amount cannot be negative.")
        End If

        _balance += amount
    End Sub

    ' Withdraws money from the account
    Public Sub Withdraw(amount As Decimal)
        If amount < 0D Then
            Throw New ArgumentException("Withdrawal amount cannot be negative.")
        End If

        If amount > _balance Then
            Throw New InvalidOperationException("Insufficient funds. Overdraft is not allowed.")
        End If

        _balance -= amount
    End Sub

End Class
