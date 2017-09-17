namespace EFData
{
    using Common;
    using System;
    using System.Collections.Generic;
    using System.Data.Entity;
    using System.Linq;

    public class EFBase : DbContext
    {
        public EFBase()
            : base("name=EFBase")
        {
            Database.SetInitializer(new StructureContextInitializer());
        }

        public virtual DbSet<Email> MyEmails { get; set; }
    }

    class StructureContextInitializer : DropCreateDatabaseIfModelChanges<EFBase>
    {
        protected override void Seed(EFBase context)
        {

            context.MyEmails.AddRange(new List<Email>()
            {
                new Email{ Name = "kobernicyri@mail.ru", Value="smtp.mail.ru", Port=25 },
                new Email{ Name = "kobernicbeljr@gmail.com", Value="smtp.gmail.com", Port=25},
                new Email{ Name = "kroshechka07@mail.ru", Value="smtp.mail.ru", Port=25},
            });

            context.SaveChanges();
        }
    }
}