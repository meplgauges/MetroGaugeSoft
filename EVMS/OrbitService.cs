using Solartron.Orbit3;

public class OrbitService : IDisposable
{
    private OrbitServer? _orbServer;
    private OrbitNetwork? _orbNet;
    private OrbitNetworks? _orbNets;
    private OrbitModules? _orbModules;
    private readonly Dictionary<string, dynamic> _orbModuleById = new();

    public IReadOnlyDictionary<string, dynamic> ModulesById => _orbModuleById;

    public int NetworkCount => _orbNets?.Count ?? 0;
    public int ModuleCount => _orbModules?.Count ?? 0;

    public bool IsConnected => _orbServer?.Connected ?? false;

    public async Task<bool> ConnectAsync()
    {
        try
        {
            if (_orbServer == null)
                _orbServer = new OrbitServer();

            if (!_orbServer.Connected)
                _orbServer.Connect();

            if (!_orbServer.Connected)
                return false;

            _orbNets = _orbServer.Networks;
            if (_orbNets == null || _orbNets.Count == 0)
                return false;

            _orbNet = _orbNets[0];
            if (_orbNet == null)
                return false;

            _orbModules = _orbNet.Modules;
            if (_orbModules == null || _orbModules.Count == 0)
            {
                _orbModules?.FindHotswapped();
                await Task.Delay(1000);
                if (_orbModules?.Count == 0)
                    return false;
            }

            _orbModuleById.Clear();
            for (int i = 0; i < _orbModules.Count; i++)
            {
                var module = _orbModules[i];
                if (module == null) continue;
                string id = module.ModuleID;
                if (!string.IsNullOrEmpty(id))
                    _orbModuleById[id] = module;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    public void Disconnect()
    {
        try
        {
            if (_orbServer != null && _orbServer.Connected)
                _orbServer.Disconnect();
        }
        catch
        {
            // ignore exceptions on disconnect
        }
    }

    public bool IsModuleConnected(string moduleId)
    {
        return !string.IsNullOrEmpty(moduleId) && _orbModuleById.ContainsKey(moduleId);
    }

    public dynamic? GetModuleById(string moduleId)
    {
        if (IsModuleConnected(moduleId))
            return _orbModuleById[moduleId];
        return null;
    }

    // New method to refresh modules (hot-swap detection)
    public void RefreshModules()
    {
        _orbModules?.FindHotswapped();
    }

    // New method to get all connected Module IDs
    public List<string> GetConnectedModuleIds()
    {
        var connectedIds = new List<string>();
        if (_orbModules == null)
            return connectedIds;
        for (int i = 0; i < _orbModules.Count; i++)
        {
            var mod = _orbModules[i];
            dynamic dmod = mod;
            connectedIds.Add(dmod.ModuleID);
        }
        return connectedIds;
    }

    // New method to clear all modules
    public void ClearModules()
    {
        _orbModules?.ClearModules();
    }

    // New method to notify addition of a module (blocking call)
    public bool NotifyAddModule()
    {
        if (_orbModules == null)
            return false;
        return _orbModules.NotifyAddModule();
    }

    // New method to get the last module in the collection
    public dynamic GetLastModule()
    {
        if (_orbModules == null || _orbModules.Count == 0)
            return null;
        return _orbModules[_orbModules.Count - 1];
    }

    public void Dispose()
    {
        Disconnect();
        _orbServer = null;
        _orbNet = null;
        _orbNets = null;
        _orbModules = null;
        _orbModuleById.Clear();
    }
}
