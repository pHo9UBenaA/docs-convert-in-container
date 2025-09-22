using System;

namespace SharedXmlToJsonl.Processing;

public interface IBufferManager : IDisposable
{
    byte[] RentBuffer(int minimumSize = 4096);
    void ReturnBuffer(byte[] buffer);
    char[] RentCharBuffer(int minimumSize = 4096);
    void ReturnCharBuffer(char[] buffer);
}