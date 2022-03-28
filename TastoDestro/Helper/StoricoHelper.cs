using Dapper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TastoDestro.Model;

namespace TastoDestro.Helper
{
    class StoricoHelper :  SqliteDB,IStorico
    {
        Storico IStorico.GetStorico(int id)
        {
            if (!File.Exists(DbFile)) return null;

            using (var cnn = SimpleDbConnection())
            {
                cnn.Open();
                Storico result = cnn.Query<Storico>(
                    @"SELECT Id, NomeFile,DataSalvataggio
                    FROM Storico
                    WHERE Id = @id", new { id }).FirstOrDefault();
                return result;
            }
        }

        public void SalvaStorico(Storico storico)
        {
            if (!File.Exists(DbFile))
            {
                CreateDatabase();
            }

            using (var cnn = SimpleDbConnection())
            {
                cnn.Open();
                storico.ID = cnn.Query<int>(
                    @"INSERT INTO Storico
                    ( NomeFile, DataSalvataggio ) VALUES 
                    ( @NomeFile, @DataSalvataggio);
                    select last_insert_rowid()", storico).First();
            }
        }



        private static void CreateDatabase()
        {
            using (var cnn = SimpleDbConnection())
            {
                cnn.Open();
                cnn.Execute(
                    @"create table Storico
                      (
                         ID                                  integer primary key AUTOINCREMENT,
                         NomeFile                            text not null,
                         DataSalvataggio                            text not null
                      )");
            }
        }

    }
}
