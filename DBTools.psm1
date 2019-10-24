function New-DBConnection {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string]$s,
        [Parameter(Mandatory = $false)]
        [string]$u,
        [Parameter(Mandatory = $false)]
        [string]$p,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Oracle", "MSSQL", "SqlServer")]
        [string]$v,
        [Parameter(Mandatory = $false)]
        [string]$d,
        [switch]$SkipTest
    )

    # Assume encrypted password equals username - possibility open to override this
    if ($p -like "TripleDes,*") {$p = $u}

    Switch ($v) {
        Oracle {
            if (!$u) {$u = 'READ_ONLY'}
            if (!$p -and $u) {$p = $u}
            $connectionString = "Data Source=$s;User ID=$u;password=$p;Provider=OraOLEDB.Oracle;"
        }

        {$_ -in "MSSQL", "SqlServer"} {
            if (!$p -and $u) {$p = $u}
            if ($d) {
                if (!$u) {$connectionString = "Data Source=$s;Initial Catalog=$d;Integrated Security=SSPI;Provider=SQLNCLI11;"}
                else {$connectionString = "Data Source=$s;Initial Catalog=$d;User Id=$u;Password=$p;Provider=SQLNCLI11;"}
            }
            else {$connectionString = "Data Source=$s;Initial Catalog=$u;User Id=$u;Password=$p;Provider=SQLNCLI11;"}
        }

        Default {
            $msg = "DB Vendor '$v' not recognized. Accepted values are 'Oracle' or 'MSSQL'."
            Throw $msg
        }
    }

    Try {
        if (!$SkipTest.IsPresent) {
            Write-Verbose "Testing connection $connectionString"
            $OLEDBConn = New-Object System.Data.OleDb.OleDbConnection
            $OLEDBConn.ConnectionString = $connectionString
            $OLEDBConn.Open()
            $OLEDBConn.Close()
            Write-Verbose "Test connection successful"
        }

        return $connectionString
    }
    Catch {
        $ex = $_.Exception
        $msg = "Error testing connection '$connectionString' - $($ex.Message)"
        Throw $msg
    }

} #end function New-DBConnection
New-Alias -Name ndbc -Value New-DBConnection

function Execute-Query {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string]$ConnectionString,
        [Parameter(Mandatory = $true)]
        [string]$SQL,
        [int]$CommandTimeout = 1000
    )

    Write-Debug "Execute-Query: $SQL;"
    Try {
        $OLEDBConn = New-Object System.Data.OleDb.OleDbConnection
        $OLEDBConn.ConnectionString = $ConnectionString
        $OLEDBConn.Open()
        $command = New-Object System.Data.OleDb.OleDbCommand($SQL, $OLEDBConn)
        $command.CommandTimeout = $CommandTimeout
        $da = New-Object System.Data.OleDb.OleDbDataAdapter($command)
        $dt = New-Object System.Data.datatable
        [void]$da.Fill($dt)
        $OLEDBConn.Close()

        return $dt
    }
    Catch {
        $ex = $_.Exception
        $msg = "Error executing '$SQL' - $($ex.Message)"
        Throw $msg
    }
} #end function Execute-Query
New-Alias -Name eq -Value Execute-Query

function Execute-NonQuery {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string]$ConnectionString,
        [Parameter(Mandatory = $true)]
        [string]$SQL,
        [int]$CommandTimeout = 1000
    )

    Write-Debug "Execute-NonQuery: $SQL;"
    Try {
        $OLEDBConn = New-Object System.Data.OleDb.OleDbConnection
        $OLEDBConn.ConnectionString = $ConnectionString
        $OLEDBConn.Open()
        $command = New-Object System.Data.OleDb.OleDbCommand($SQL, $OLEDBConn)
        $command.CommandTimeout = $CommandTimeout
        $numRowsAffected = $command.ExecuteNonQuery()
        $OLEDBConn.Close()

        return $numRowsAffected
    }
    Catch {
        $ex = $_.Exception
        $msg = "Error executing '$SQL' - $($ex.Message)"
        Throw $msg
    }
} #end function Execute-NonQuery
New-Alias -Name enq -Value Execute-NonQuery

Export-ModuleMember -Function New-DBConnection, Execute-Query, Execute-NonQuery -Alias ndbc, eq, enq
