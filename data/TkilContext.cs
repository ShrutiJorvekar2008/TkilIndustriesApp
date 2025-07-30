using Microsoft.EntityFrameworkCore;
using TkilIndustriesApp.Models;

namespace TkilIndustriesApp.Data
{
    public class TkilContext : DbContext
    {
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ThirdPartyData>()
                .HasKey(t => new { t.Employee8ID, t.Email });

            base.OnModelCreating(modelBuilder);
        }
        public TkilContext(DbContextOptions<TkilContext> options) : base(options) { }

        public DbSet<ThirdPartyData> IT_ThirdPartyRecords { get; set; }
        public DbSet<IT> IT_ITRecords { get; set; }
        public DbSet<LicenseMaster> IT_LicenseMaster {  get; set; }
    }
}
