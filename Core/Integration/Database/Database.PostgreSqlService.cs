namespace SKSB_App.Core.Integration.Database;

using Dapper;
using Npgsql;
using SKSB_App.Core.Diagnostics;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

public class PostgreSqlService
{
    public const string DefaultConnectionString = "Host=192.168.0.20;Port=5432;Database=sksb_db;Username=sksb_server;Password=sksb7811;Timeout=5;";

    private readonly string _connectionString;

    public PostgreSqlService(string? connectionString = null)
    {
        _connectionString = connectionString ?? DefaultConnectionString;
    }

    private NpgsqlConnection CreateConnection()
    {
        ServerEnvironment.GuardConnection();
        return new NpgsqlConnection(_connectionString);
    }

    /// <summary>Runs a server call; a network-level failure flips the app into Offline Mode before rethrowing.</summary>
    private static T Guarded<T>(Func<T> action)
    {
        try
        {
            return action();
        }
        catch (Exception ex)
        {
            if (ServerEnvironment.IsConnectionException(ex)) ServerEnvironment.EnterOfflineMode(ex.Message);
            throw;
        }
    }

    public IEnumerable<T> Query<T>(string sql, object? param = null)
    {
        return Guarded(() =>
        {
            using var db = CreateConnection();
            return db.Query<T>(sql, param).ToList();
        });
    }

    public void ExecuteTransaction<T>(string sql, IEnumerable<T> records)
    {
        Guarded(() =>
        {
            using var db = CreateConnection();
            db.Open();
            using var trans = db.BeginTransaction();
            db.Execute(sql, records, transaction: trans);
            trans.Commit();
            return 0;
        });
    }

    public void ExecuteInTransaction(Action<IDbConnection, IDbTransaction> action)
    {
        Guarded(() =>
        {
            using var db = CreateConnection();
            db.Open();
            using var trans = db.BeginTransaction();
            try
            {
                action(db, trans);
                trans.Commit();
            }
            catch (PostgresException pgEx)
            {
                trans.Rollback();

                var errSb = new StringBuilder();
                errSb.AppendLine("==========================================================================");
                errSb.AppendLine($" POSTGRESQL TRANSACTION ERROR [{pgEx.SqlState}]");
                errSb.AppendLine("==========================================================================");
                errSb.AppendLine($"Message  : {pgEx.Message}");
                errSb.AppendLine($"Position : {pgEx.Position}");
                errSb.AppendLine($"Table    : {pgEx.TableName}");
                errSb.AppendLine($"Column   : {pgEx.ColumnName}");
                errSb.AppendLine($"Where    : {pgEx.Where}");
                errSb.AppendLine($"Routine  : {pgEx.Routine}");

                AppLogger.Error(errSb.ToString(), pgEx);
                throw;
            }
            catch (Exception ex)
            {
                try { trans.Rollback(); } catch { }
                AppLogger.Error("ExecuteInTransaction general error", ex);
                throw;
            }
            return 0;
        });
    }

    public T? ExecuteScalar<T>(string sql, object? param = null)
    {
        return Guarded(() =>
        {
            using var db = CreateConnection();
            db.Open();
            return db.ExecuteScalar<T>(sql, param);
        });
    }
}
