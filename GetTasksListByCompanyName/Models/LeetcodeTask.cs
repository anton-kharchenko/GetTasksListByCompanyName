using Newtonsoft.Json;

namespace GetTasksListByCompanyName.Models;

public class LeetcodeTask
{
    [JsonProperty("num_occur")]
    public int NumOccur { get; set; }

    [JsonProperty("company")]
    public string Company { get; set; }
}