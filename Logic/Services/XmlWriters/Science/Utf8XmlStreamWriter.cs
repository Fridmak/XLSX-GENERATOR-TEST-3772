using System.Runtime.InteropServices;
using System.Text;

namespace Analitics6400.Logic.Services.XmlWriters.Science;

internal sealed class Utf8XmlStreamWriter : IDisposable
{
    private readonly Stream _stream;
    private readonly byte[] _buffer;
    private int _position;
    private bool _disposed;

    public Utf8XmlStreamWriter(Stream stream, int bufferSize = 64 * 1024)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _buffer = new byte[bufferSize];

        var c = stream is MemoryStream;

        WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    }

    public void WriteStart(string tag, params string[] attributes)
    {
        if (attributes.Length % 2 != 0)
            throw new ArgumentException("Attributes must be in pairs: name, value");

        WriteByte('<');
        WriteRaw(tag);

        for (int i = 0; i < attributes.Length; i += 2)
        {
            WriteByte(' ');
            WriteRaw(attributes[i]);
            WriteRaw("=\"");
            WriteEscaped(attributes[i + 1].AsSpan());
            WriteByte('"');
        }

        WriteByte('>');
    }

    public void WriteStart(string tag)
    {
        WriteByte('<');
        WriteRaw(tag);
        WriteByte('>');
    }

    public void WriteEnd(string tag)
    {
        WriteRaw("</");
        WriteRaw(tag);
        WriteByte('>');
    }

    public void WriteText(string text)
    {
        WriteEscaped(text.AsSpan());
    }

    public void WriteText(ReadOnlySpan<char> text)
    {
        WriteEscaped(text);
    }

    public void WriteEmptyElement(string tag, params string[] attributes)
    {
        if (attributes.Length % 2 != 0)
            throw new ArgumentException("Attributes must be in pairs: name, value");

        WriteByte('<');
        WriteRaw(tag);

        for (int i = 0; i < attributes.Length; i += 2)
        {
            WriteByte(' ');
            WriteRaw(attributes[i]);
            WriteRaw("=\"");
            WriteEscaped(attributes[i + 1].AsSpan());
            WriteByte('"');
        }

        WriteRaw("/>");
    }

    private void WriteEscaped(ReadOnlySpan<char> text)
    {
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '<':
                    WriteRaw("&lt;");
                    break;
                case '>':
                    WriteRaw("&gt;");
                    break;
                case '&':
                    WriteRaw("&amp;");
                    break;
                case '"':
                    WriteRaw("&quot;");
                    break;
                case '\'':
                    WriteRaw("&apos;");
                    break;
                default:
                    if (ch >= 0x20 || ch == '\t' || ch == '\n' || ch == '\r')
                    {
                        WriteUtf8Char(ch);
                    }
                    break;
            }
        }
    }

    private void WriteUtf8Char(char ch)
    {
        Span<byte> temp = stackalloc byte[4];
        int bytesWritten = Encoding.UTF8.GetBytes(MemoryMarshal.CreateReadOnlySpan(ref MemoryMarshal.GetReference(
            MemoryMarshal.CreateSpan(ref ch, 1)), 1), temp);

        Write(temp.Slice(0, bytesWritten));
    }

    private void Write(ReadOnlySpan<byte> data)
    {
        if (_position + data.Length > _buffer.Length)
            Flush();

        data.CopyTo(_buffer.AsSpan(_position));
        _position += data.Length;
    }

    private void WriteRaw(string text)
    {
        Write(Encoding.UTF8.GetBytes(text));
    }

    private void WriteByte(char c)
    {
        if (_position >= _buffer.Length)
            Flush();

        _buffer[_position++] = (byte)c;
    }

    public void Flush()
    {
        if (_position > 0)
        {
            _stream.Write(_buffer, 0, _position);
            _position = 0;
        }
        _stream.Flush();
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            Flush();
            _disposed = true;
        }
    }
}