public class PlayerInfo
{
    // 序号
    public string ID { get; set; }

    // 姓名
    public string Name { get; set; }

    // 学号
    public string Number { get; set; }

    // 班级
    public string ClassInfo { get; set; }

    /// 手机号
    public string Mobile { get; set; }

    public PlayerInfo(string id, string name, string classInfo, string number, string mobile)
    {
        this.ID = id;
        this.Name = name;
        this.Number = number;
        this.ClassInfo = classInfo;
        this.Mobile = mobile;
    }
}