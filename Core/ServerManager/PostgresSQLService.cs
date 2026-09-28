namespace SKSB_App.Core.ServerManager;

using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Dapper;
using Npgsql;

public class PostgreSqlService
{
    private readonly string _connectionString;

    public PostgreSqlService(string connectionString = null)
    {
        // Centralised connection string fallback
        _connectionString = connectionString ??
            "Host=192.168.0.20;Port=5432;Database=sksb_db;Username=sksb_server;Password=sksb7811;";
    }

    private NpgsqlConnection CreateConnection() => new NpgsqlConnection(_connectionString);

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
}