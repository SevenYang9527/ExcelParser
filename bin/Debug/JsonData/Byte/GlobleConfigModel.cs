using MemoryPack;
using System.Collections.Generic;
using UnityEngine;

[MemoryPackable]
public partial class GlobleConfigConfigModel
{
    public Dictionary<string, GlobleConfig> GlobleConfig;
    public GlobleConfig GetValue(string key)
    {
        if (GlobleConfig.TryGetValue(key, out GlobleConfig value))
            return value;
        Debug.LogError($"{nameof(GlobleConfig)}未查询到key：{key}");
        return null;
    }
}
[MemoryPackable]
public partial class GlobleConfigConfig
{
    public string ID;
    public int Const1;
    public string Const2;
    public float Const3;
    public int[] Param1;
    public string[] Param2;

}
