using MemoryPack;
using System.Collections.Generic;
using UnityEngine;

[MemoryPackable]
public partial class PropConfigConfigConfigModel
{
    public Dictionary<string, PropConfigConfig> PropConfigConfig;
    public PropConfigConfig GetValue(string key)
    {
        if (PropConfigConfig.TryGetValue(key, out PropConfigConfig value))
            return value;
        Debug.LogError($"{nameof(PropConfigConfig)}未查询到key：{key}");
        return null;
    }
}
[MemoryPackable]
public partial class PropConfigConfigConfig
{
    public string ID;
    public string Name;
    public string Desc;
    public string IconName;
    public int Init;
    public int Unlock;
    public int GetByAd;
    public int Pass;
    public int Rate;

}
