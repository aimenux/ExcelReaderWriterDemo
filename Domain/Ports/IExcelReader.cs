using Domain.Models;

namespace Domain.Ports
{
    public interface IExcelReader
    {
        Stock Read(string file);
    }
}
