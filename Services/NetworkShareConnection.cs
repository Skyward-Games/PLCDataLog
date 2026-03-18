using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PLCDataLog.Services;

public sealed class NetworkShareConnection : IDisposable
{
    private readonly string? _remoteName;
    private readonly bool _connected;

    private NetworkShareConnection(string? remoteName, bool connected)
    {
        _remoteName = remoteName;
        _connected = connected;
    }

    public static NetworkShareConnection ConnectIfNeeded(string folderPath, string? username, string? password)
    {
        if (string.IsNullOrWhiteSpace(folderPath))
            return new NetworkShareConnection(null, false);

        if (!folderPath.StartsWith("\\\\", StringComparison.OrdinalIgnoreCase))
            return new NetworkShareConnection(null, false);

        if (string.IsNullOrWhiteSpace(username))
            return new NetworkShareConnection(null, false);

        var remoteName = GetShareRoot(folderPath);
        if (string.IsNullOrWhiteSpace(remoteName))
            return new NetworkShareConnection(null, false);

        var nr = new NETRESOURCE
        {
            dwType = 1,
            lpRemoteName = remoteName,
        };

        var result = WNetAddConnection2(ref nr, password, username, 0);
        if (result != 0 && result != 1219)
            throw new Win32Exception(result);

        // 1219 = já existe conexão com credenciais diferentes. Nesse caso não desconectamos no Dispose.
        var createdByThisInstance = result == 0;
        return new NetworkShareConnection(remoteName, createdByThisInstance);
    }

    public void Dispose()
    {
        if (!_connected || string.IsNullOrWhiteSpace(_remoteName))
            return;

        _ = WNetCancelConnection2(_remoteName, 0, true);
    }

    private static string? GetShareRoot(string uncPath)
    {
        var trimmed = uncPath.Trim().TrimEnd('\\');
        var parts = trimmed.Split('\\', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2)
            return null;

        return $"\\\\{parts[0]}\\{parts[1]}";
    }

    [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
    private static extern int WNetAddConnection2(ref NETRESOURCE lpNetResource, string? lpPassword, string? lpUserName, int dwFlags);

    [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
    private static extern int WNetCancelConnection2(string lpName, int dwFlags, bool fForce);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NETRESOURCE
    {
        public int dwScope;
        public int dwType;
        public int dwDisplayType;
        public int dwUsage;
        public string? lpLocalName;
        public string? lpRemoteName;
        public string? lpComment;
        public string? lpProvider;
    }
}
