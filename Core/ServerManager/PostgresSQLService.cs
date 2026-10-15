namespace SKSB_App.Core.ServerManager;

using System;
using System.Collections.Generic;
using System.Data;
using Dapper;
using Npgsql;

public static class ServerEnvironment
{
#if DEBUG
    // Default is strictly FALSE. Only the Ribbon button can turn this on.
    private static bool _simulateOffline = false;
#else
    private static readonly bool _simulateOffline = false;
#endif

    public static bool SimulateOffline
    {
        get => _simulateOffline;
        set
        {
#if DEBUG
            _simulateOffline = value;
#else
            throw new InvalidOperationException("Offline simulation is disabled in Release builds.");
#endif
        }
    }

    public static void GuardConnection()
    {
        if (SimulateOffline)
        {
            throw new TimeoutException("[Simulated Offline Mode Active] Connection to PostgreSQL server blocked by user switch.");
        }
    }
}

public class PostgreSqlService
{
    private readonly string _connectionString;

    public PostgreSqlService(string connectionString = null)
    {
        _connectionString = connectionString ??
            "Host=192.168.0.20;Port=5432;Database=sksb_db;Username=sksb_server;Password=sksb7811;";
    }

    private NpgsqlConnection CreateConnection()
    {
        ServerEnvironment.GuardConnection();
        return new NpgsqlConnection(_connectionString);
    }

    public IEnumerable<T> Query<T>(string sql, object param = null)
    {
        using var db = CreateConnection();
        return db.Query<T>(sql, param);
    }

    public int Execute(string sql, object param = null)
    {
        using var db = CreateConnection();
        return db.Execute(sql, param);
    }

    public void ExecuteTransaction<T>(string sql, IEnumerable<T> records)
    {
        using var db = CreateConnection();
        db.Open();
        using var trans = db.BeginTransaction();
        db.Execute(sql, records, transaction: trans);
        trans.Commit();
    }

    public void ExecuteInTransaction(Action<IDbConnection, IDbTransaction> action)
    {
        using var db = CreateConnection();
        db.Open();
        using var trans = db.BeginTransaction();
        try
        {
            action(db, trans);
            trans.Commit();
        }
        catch
        {
            trans.Rollback();
            throw;
        }
    }

    public T? ExecuteScalar<T>(string sql, object? param = null)
    {
        using var db = CreateConnection();
        db.Open();
        return db.ExecuteScalar<T>(sql, param);
    }
}