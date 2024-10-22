## A Excel VBA Game

**Chris Wang github@bellowaverage**

一、项目概述

-   这是一款以上海大学嘉定校区缩略图为地图的俯视角探索类游戏。游戏中，"\*"代表玩家可以自由行动的道路；黑色填充的色块代表玩家；大写英文字母代表可以探索的单元格（地点），只需要操纵色块移动到单元格上即可触发；黑色圆圈圈起来的单元格（H5和J12分别)代表体育考点和文化课考点。玩家移动将消耗行动点数，通过地图探索可以增加&减少能力数值。点击重置游戏可以随时重玩&结束当前游戏（辍学&毕业），并获得成绩（最终平均绩点）。最终成绩和游戏结束时的玩家能力数值，搜索到的物品、称号相关，与剩余行动点数，道具使用次数无关)
    特别注意：请不要在游戏过程中随意改变选中的单元格；请将右下角屏幕缩放调至40%左右。

-   游戏要求玩家使用有限的行动点数，取得最高的玩家能力数值。玩家最后的得分取决于玩家最终重置游戏时的游戏数值。最高分是4.0，最低分是零。玩家只有在未使用过商城道具的前提下，才能取得得分。

-   ![](./media/media/image2.tmp)游戏界面（sheet1）

二、项目设计初衷和意义

-   模拟现代游戏运营的基本方式，建立数据模型；为游戏数值设计师提供决策依据。现代游戏的运营在本项目被简化为前后端两部分，即游戏程序（sheet1）和后台数据分析（sheet2-4）。

-   游戏数据分析的目的是为游戏数值和玩法策划提供依据，兼顾游戏营利性和可玩性的平衡。

三、项目设计思路

> 游戏前台，即游戏页面（sheet1），包含：

-   控制台（上下左右控制玩家移动的按钮，和地图进行互动的按钮，结束&重置游戏的按钮）

-   信息栏，即用于浏览玩家目前收集到的物品（装备）情况，玩家能力数值，获得过的称号和奖励等游戏情报）

-   商城，用于与玩家进行互动消费，实现盈利目的。

-   游戏显示区域，即玩家（用色块表示），道路（用"\*"标记），地图背景，以及其他地图元素。

> 游戏后台：

-   记录下玩家每次游玩的游戏数据，并进行分析。会被记录的游戏数据，包括游戏结束时玩家数值，物品，最终游戏成绩，对游戏体验的评价，在游戏过程中做出的选择，使用商城的次数，购买道具的次数，总消费等。并基于玩家游玩数据分析游戏数值设定是否合理，是否兼顾了盈利性和可玩性，为游戏数值设计师的策划和决策提供参考和依据。

四、核心代码：移动和探索程序

Sub testw()

If Range(\"ae11\").Value \> 0 Then '检验玩家是否还有剩余行动点数

originalcolumn = ActiveCell.Column
\'将此时活动单元格的列用变量originalcolumn记录下来

row = ActiveCell.Row \'将此时活动单元格的行用变量row记录下来

row = row - 1 \'向上移动1格即目标单元格是原所在单元格行的大小减1

If Cells (activecell.column , row).Value = \" \*\" Then
\'检验玩家是否在道路上行走

Worksheets(\"sheet1\").Cells(activecell.column , row).Interior.Color =
vbred \'为目标单元格上色

row = row + 1 \'将行的大小还原

Worksheets(\"sheet1\").Cells(activecell.column , row).Clear
\'清空原单元格

Worksheets(\"sheet1\").Cells(activecell.column , row).Value = \" \*\"
\'给原单元格赋值"\*"

With Worksheets(\"sheet1\").Cells(activecell.column , row).Font
\'赋予原单元格格式（不同的\*号颜色以区别走过的道路）

.Size = 36

.Color = -1677696

.Bold = True

End With

Cells(row - 1, originalcolumn).Select
\'选中当前单元格，确保下次向上移动命令被执行时，活动单元格的状态是最新的

Range(\"ae11\").Value = Range(\"ae11\").Value - 1 \'剩余行动点数减1

Else

Select Case Range(colum & lie).Value

Case \"I\"

If IsEmpty(Sheet1.Range(\"ab17\")) = False Then

MsgBox \"不能重复探索!\"

Else

If MsgBox(\"确定要探索该区域吗？\", vbYesNo + vbInformation,
\"游戏提示\") = vbYes Then

MsgBox
\"欢迎来到5号男生宿舍楼！恭喜你获得了物品：《5号男生宿舍楼名单》！获得奖励：体力+5
智力+2 品德+3 审美+3 劳动能力+5 行动点数+5\", vbOKOnly, \"游戏提示\"

Range(\"ab9\").Value = Range(\"ab9\").Value + 3

Range(\"ad9\").Value = Range(\"ad9\").Value + 2

Range(\"af9\").Value = Range(\"af9\").Value + 5

Range(\"ah9\").Value = Range(\"ah9\").Value + 3

Range(\"aj9\").Value = Range(\"aj9\").Value + 5

Range(\"ae11\").Value = Range(\"ae11\").Value + 5

Range(\"ab17\").Value = \"《5号男生宿舍楼名单》\"

End If

End If

Case \"F\"

If IsEmpty(Sheet1.Range(\"ad17\")) = False Then

MsgBox \"不能重复探索!\"

Else

If MsgBox(\"确定要探索该区域吗？\", vbYesNo + vbInformation,
\"游戏提示\") = vbYes Then

If
MsgBox(\"欢迎来到悉商一食堂！恭喜你获得了物品：曹轶骏的饭卡！你是否要将饭卡归还给曹轶骏？\",
vbYesNo + vbInformation, \"游戏提示\") = vbYes Then

MsgBox
\"年轻人品德高尚！恭喜你获得了物品：《助人为乐奖状》！获得奖励：体力+0
智力+0 品德+20 审美+0 劳动能力+0 行动点数+5\", vbOKOnly, \"游戏提示\"

Range(\"ab9\").Value = Range(\"ab9\").Value + 20

Range(\"ad9\").Value = Range(\"ad9\").Value + 0

Range(\"af9\").Value = Range(\"af9\").Value + 0

Range(\"ah9\").Value = Range(\"ah9\").Value + 3

Range(\"aj9\").Value = Range(\"aj9\").Value + 5

Range(\"ae11\").Value = Range(\"ae11\").Value + 5

Range(\"ad17\").Value = \"《助人为乐奖状》\"

Else: MsgBox \"年轻人不讲武德！获得：体力+0 智力+5 品德-20 审美+0
劳动能力+0 行动点数-5\", vbOKOnly, \"游戏提示\"

Range(\"ab9\").Value = Range(\"ab9\").Value - 20

Range(\"ad9\").Value = Range(\"ad9\").Value + 5

Range(\"af9\").Value = Range(\"af9\").Value + 0

Range(\"ah9\").Value = Range(\"ah9\").Value + 0

Range(\"aj9\").Value = Range(\"aj9\").Value + 0

Range(\"ae11\").Value = Range(\"ae11\").Value - 5

Range(\"ad17\").Value = \"不讲武德\"

End If

End If

End If

Case \"C\"

If IsEmpty(Sheet1.Range(\"af17\")) = False Then

MsgBox \"不能重复探索!\"

Else

If MsgBox(\"确定要探索该区域吗？\", vbYesNo + vbInformation,
\"游戏提示\") = vbYes Then

If
MsgBox(\"欢迎来到文达楼！这里是PATHWAY学生的上课区域。一个过路的同学向你提出了这样的问题：do
you think your mom is closer to you than your dad?\", vbYesNo +
vbInformation, \"游戏提示\") = vbYes Then

MsgBox \"回答错误！因为：dad is
farther!恭喜你获得了称号：误人子弟！获得：体力+0 智力-10 品德-10 审美+0
劳动能力+0 行动点数+0\", vbOKOnly, \"游戏提示\"

Range(\"ab9\").Value = Range(\"ab9\").Value - 10

Range(\"ad9\").Value = Range(\"ad9\").Value - 10

Range(\"af9\").Value = Range(\"af9\").Value + 0

Range(\"ah9\").Value = Range(\"ah9\").Value + 0

Range(\"aj9\").Value = Range(\"aj9\").Value + 0

Range(\"ae11\").Value = Range(\"ae11\").Value + 0

Range(\"af17\").Value = \"误人子弟\"

Else: MsgBox \"回答正确！因为：dad is farther！获得称号：Bless from
Doris 获得奖励：体力+0 智力+10 品德+10 审美+0 劳动能力+0 行动点数+5\",
vbOKOnly, \"游戏提示\"

Range(\"ab9\").Value = Range(\"ab9\").Value + 10

Range(\"ad9\").Value = Range(\"ad9\").Value + 10

Range(\"af9\").Value = Range(\"af9\").Value + 0

Range(\"ah9\").Value = Range(\"ah9\").Value + 0

Range(\"aj9\").Value = Range(\"aj9\").Value + 0

Range(\"ae11\").Value = Range(\"ae11\").Value + 5

Range(\"af17\").Value = \"Bless from Doris\"

End If

End If

End If

Case \"E\"

If IsEmpty(Sheet1.Range(\"ah17\")) = False Then

MsgBox \"不能重复探索!\"

Else

If MsgBox(\"确定要探索该区域吗？\", vbYesNo + vbInformation,
\"游戏提示\") = vbYes Then

If MsgBox(\"欢迎来到足球场！获得物品：足球。是否进行1000米跑步训练？\",
vbYesNo + vbInformation, \"游戏提示\") = vbYes Then

MsgBox \"生命在于运动，但是消耗了大量体力。获得：体力+20 智力+0 品德+0
审美+0 劳动能力+0 行动点数-5\", vbOKOnly, \"游戏提示\"

Range(\"ab9\").Value = Range(\"ab9\").Value - 0

Range(\"ad9\").Value = Range(\"ad9\").Value - 0

Range(\"af9\").Value = Range(\"af9\").Value + 20

Range(\"ah9\").Value = Range(\"ah9\").Value + 0

Range(\"aj9\").Value = Range(\"aj9\").Value + 0

Range(\"ae11\").Value = Range(\"ae11\").Value - 5

Range(\"ah17\").Value = \"足球\"

Else: MsgBox \"错失了训练的机会！但是你保存了体力。获得奖励：体力-10
智力+0 品德+0 审美+0 劳动能力+0 行动点数+5\", vbOKOnly, \"游戏提示\"

Range(\"ab9\").Value = Range(\"ab9\").Value + 0

Range(\"ad9\").Value = Range(\"ad9\").Value + 0

Range(\"af9\").Value = Range(\"af9\").Value - 10

Range(\"ah9\").Value = Range(\"ah9\").Value + 0

Range(\"aj9\").Value = Range(\"aj9\").Value + 0

Range(\"ae11\").Value = Range(\"ae11\").Value + 5

Range(\"ah17\").Value = \"足球\"

End If

End If

End If

Case \"D\"

If IsEmpty(Sheet1.Range(\"aj17\")) = False Then

MsgBox \"不能重复探索!\"

Else

If MsgBox(\"确定要探索该区域吗？\", vbYesNo + vbInformation,
\"游戏提示\") = vbYes Then

UserForm1.Show

End If

End If

Case Else

MsgBox \"你不能在道路以外的地图区域上移动\", vbOKOnly + vbInformation,
\"游戏提示\"

End Select

End If

Else: MsgBox
\"您已经耗尽行动点数，请点击重置游戏按钮来结束游戏；如果这不是您第一次游戏或您想充分探索，请至商城购买行动点数。\"

End If

End Sub

五、其他代码：考试，商城，数据采集，重置游戏等模块。详见VBA程序。

六、数据分析

-   玩家探索兴趣分布分析图表

-   玩家在"文德楼事件"中选择倾向分析图表

-   等详见VBA程序
