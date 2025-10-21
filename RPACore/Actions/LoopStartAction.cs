using System;
using System.Threading.Tasks;

namespace RPACore.Actions;

/// <summary>
/// ループ開始アクション
/// </summary>
public class LoopStartAction : ActionBase
{
    /// <summary>コメント（ループの説明）</summary>
    public string Comment { get; set; } = string.Empty;

    public override string Name => "ループ開始";

    public override string Description
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(Comment))
            {
                return $"ループ開始: {Comment}";
            }
            return "ループ開始";
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        await Task.CompletedTask;

        LogInfo($"ループを開始します");

        // ループ開始マーカー（実際のループ制御はLoopEndActionで行う）
        return true;
    }
}
