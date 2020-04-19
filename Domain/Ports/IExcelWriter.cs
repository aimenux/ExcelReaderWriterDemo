using Domain.Models;

namespace Domain.Ports
{
    public interface IExcelWriter
    {
        bool Write(string file, Stock stock);
    }
}