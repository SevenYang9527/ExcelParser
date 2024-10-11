using MemoryPack;
using System.Collections.Generic;
using UnityEngine;

[MemoryPackable]
public partial class ClubActiveTaskConfigConfigModel
{
    public Dictionary<string, ClubActiveTaskConfig> ClubActiveTaskConfig;
    public ClubActiveTaskConfig GetValue(string key)
    {
        if (ClubActiveTaskConfig.TryGetValue(key, out ClubActiveTaskConfig value))
            return value;
        Debug.LogError($"{nameof(ClubActiveTaskConfig)}未查询到key：{key}");
        return null;
    }
}
[MemoryPackable]
public partial class ClubActiveTaskConfigConfig
{
    public string ID;
    public string TaskDesc;
    public int RequireNum;
    public string Reward;

}
