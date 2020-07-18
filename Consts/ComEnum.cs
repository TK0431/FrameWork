using FrameWork.Utility;
using System;
using System.ComponentModel;

namespace FrameWork.Consts
{
    /// <summary>
    /// 控件类型
    /// </summary>
    public enum HwndItemType
    {
        [Description("应用程序")]
        Exe,
        [Description("组控件")]
        Group,
        [Description("控件")]
        Control
    }

    /// <summary>
    /// 区分枚举
    /// </summary>
    public enum EnumType
    {
        [Value("01"), Description("画面"), English("View"), Japanese("画面")]
        Type_01,
        [Value("02"), Description("权限"), English("Level"), Japanese("権限")]
        Type_02,
        [Value("03"), Description("作业区分"), English("Work Type"), Japanese("作業区分")]
        Type_03,
        [Value("04"), Description("语言"), English("Language"), Japanese("言語")]
        Type_04,
        [Value("05"), Description("消息"), English("Message"), Japanese("メッセージ")]
        Type_05
    }

    /// <summary>
    /// 作业区分
    /// </summary>
    public enum EnumLanguage
    {
        [Value("00"), Description(""), English(""), Japanese("")]
        ALL,
        [Value("CN"), Description("中文"), English("Chinese"), Japanese("中国語")]
        CN,
        [Value("EN"), Description("英语"), English("English"), Japanese("英語")]
        EN,
        [Value("JP"), Description("日语"), English("Japanese"), Japanese("日本語")]
        JP,
    }

    /// <summary>
    /// 权限
    /// </summary>
    public enum EnumLevel
    {
        [Value("0"), Description(""), English(""), Japanese("")]
        LEVEL_0,
        [Value("1"), Description("组员"), English("Member"), Japanese("")]
        LEVEL_1,
        [Value("2"), Description("组长"), English("Leader"), Japanese("")]
        LEVEL_2,
        [Value("3"), Description("项目负责人"), English(""), Japanese("")]
        LEVEL_3,
        [Value("4"), Description("项目经理"), English(""), Japanese("")]
        LEVEL_4,
        [Value("9"), Description("超级"), English(""), Japanese("")]
        LEVEL_9,
    }

    /// <summary>
    /// 作业区分
    /// </summary>
    public enum EnumDevType
    {
        [Value("00"), Description("")]
        ALL,
        [Value("01"), Description("基本设计"), English(""), Japanese("")]
        BS,
        [Value("02"), Description("基本设计检证"), English(""), Japanese("")]
        BS_RE,
        [Value("03"), Description("详细设计"), English(""), Japanese("")]
        DL,
        [Value("04"), Description("详细设计检证"), English(""), Japanese("")]
        DL_RE,
        [Value("05"), Description("单体制造"), English(""), Japanese("")]
        CD,
        [Value("06"), Description("单体制造检证"), English(""), Japanese("")]
        CD_RE,
        [Value("07"), Description("单体测试"), English(""), Japanese("")]
        UT,
        [Value("08"), Description("单体测试检证"), English(""), Japanese("")]
        UT_RE,
        [Value("09"), Description("结合测试"), English(""), Japanese("")]
        IT,
        [Value("10"), Description("结合测试检证"), English(""), Japanese("")]
        IT_RE,
        [Value("11"), Description("综合测试"), English(""), Japanese("")]
        CT,
        [Value("12"), Description("综合测试检证"), English(""), Japanese("")]
        CT_RE,
    }

    /// <summary>
    /// 消息
    /// </summary>
    public enum EnumMessage
    {
        [Value("0000"), Description(""), English(""), Japanese("")]
        ALL,
        [Value("I001"), Description("登录成功"), English("Login success"), Japanese("登録成功")]
        I001,
        [Value("W001"), Description("登录失败"), English("Login faild"), Japanese("登録失敗")]
        W001,
    }

    public enum EnumMessageType
    {
        Info,
        Warn,
        Error,
    }

    /// <summary>
    /// 将枚举更新到数据库
    /// </summary>
    public static class UpdateEnum
    {
        public static void Update()
        {
            using (EntityUtility db = new EntityUtility())
            {
                db.DeleteAll("delete from tb_type");

                //AddDatas(db, typeof(PageEnum), EnumType.Type_01.GetValue());
                AddDatas(db, typeof(EnumLevel), EnumType.Type_02.GetValue());
                AddDatas(db, typeof(EnumDevType), EnumType.Type_03.GetValue());
                AddDatas(db, typeof(EnumLanguage), EnumType.Type_04.GetValue());
                AddDatas(db, typeof(EnumMessage), EnumType.Type_05.GetValue());
            }
        }

        private static void AddDatas(EntityUtility db, Type type, string typeValue)
        {
            //List<TB_Type> datas = new List<TB_Type>();

            //EnumLanguage back = App.Language;

            //typeof(EnumLanguage).GetList(true).ForEach(x =>
            //{
            //    App.Language = (EnumLanguage)x.Index;
            //    List<EnumItem> list = type.GetList();

            //    list.ForEach(y => datas.Add(new TB_Type()
            //    {
            //        Type = typeValue,
            //        Value = y.Value,
            //        Name = y.Description,
            //    }));
            //});

            //App.Language = back;

            //db.AddData(datas);
        }
    }
}
