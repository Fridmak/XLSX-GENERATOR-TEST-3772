namespace Analitics6400.Logic.Models;

class CountingStream : Stream
{
    public long LengthWritten { get; private set; }

    public override void Write(byte[] buffer, int offset, int count)
    {
        LengthWritten += count;
    }

    public override void WriteByte(byte value)
    {
        LengthWritten += 1;
    }

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => LengthWritten;
    public override long Position { get => LengthWritten; set => throw new NotSupportedException(); }
    public override void Flush() { }
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
}

