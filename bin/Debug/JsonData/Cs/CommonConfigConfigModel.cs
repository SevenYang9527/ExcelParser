using MemoryPack;
using System.Collections.Generic;
using UnityEngine;

[MemoryPackable]
public partial class CommonConfigConfigConfigModel
{
    public Dictionary<string, CommonConfigConfig> CommonConfigConfig;
    public CommonConfigConfig GetValue(string key)
    {
        if (CommonConfigConfig.TryGetValue(key, out CommonConfigConfig value))
            return value;
        Debug.LogError($"{nameof(CommonConfigConfig)}未查询到key：{key}");
        return null;
    }
}
[MemoryPackable]
public partial class CommonConfigConfig
{
    public string Version;
    public int VersionCode;
    public string Bulletin;
    public bool IsOpenAd;
    public string VideoId;
    public string BannerId;
    public string CustomId;
    public string MoreCustomId;
    public string InterstitialId;

}
