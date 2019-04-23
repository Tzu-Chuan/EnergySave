<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="total_LocalConditions.aspx.cs" Inherits="WebPage_total_LocalConditions" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">因地制宜 - 各縣市申請數</span></div>
            <div class="right">附加圖表 / 管理員總覽表 / 因地制宜 - 各縣市申請數</div>
        </div>

        期別：
        <select id="ddlStage" class="inputex" style="margin-top: 10px;">
            <option value="1" selected="selected">第一期</option>
            <option value="2">第二期</option>
            <option value="3">第三期</option>
        </select><br /><br />
        <div class="stripeMe">
            <table id="airlist" border="0" cellspacing="0" cellpadding="0" width="100%">
                <thead>
                    <tr>
                        <th rowspan="2" style="width:10%">縣市</th>
                        <th rowspan="2">年/月</th>
                        <th colspan="3">開飲機(台)</th>
                        <th colspan="3">溫熱水瓶(台)</th>
                        <th colspan="3">分離式冷氣(台)</th>
                        <th colspan="3">LED燈泡(顆)</th>
                        <th colspan="3">儲備型電熱水器(台)</th>
                    </tr>
                    <tr>
                        <th>規劃數</th>
                        <th>申請數</th>
                        <th>完成數</th>
                        <th>規劃數</th>
                        <th>申請數</th>
                        <th>完成數</th>
                        <th>規劃數</th>
                        <th>申請數</th>
                        <th>完成數</th>
                        <th>規劃數</th>
                        <th>申請數</th>
                        <th>完成數</th>
                        <th>規劃數</th>
                        <th>申請數</th>
                        <th>完成數</th>
                    </tr>
                </thead>
                <tbody id="tbody_list">
                    <tr class="alt">
                        <td style="text-align: center;">基隆市</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">4720.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">5140</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">550</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">10</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">3</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">臺北市</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">63610.0</td>
                        <td style="text-align: right;">20575.0</td>
                        <td style="text-align: right;">2193.0</td>
                        <td style="text-align: right;">186785</td>
                        <td style="text-align: right;">35003</td>
                        <td style="text-align: right;">2180</td>
                        <td style="text-align: right;">190787</td>
                        <td style="text-align: right;">14513</td>
                        <td style="text-align: right;">1998</td>
                        <td style="text-align: right;">197</td>
                        <td style="text-align: right;">20</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">31</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr class="alt">
                        <td style="text-align: center;">新北市</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">55300.0</td>
                        <td style="text-align: right;">2240.7</td>
                        <td style="text-align: right;">2240.7</td>
                        <td style="text-align: right;">165600</td>
                        <td style="text-align: right;">12066</td>
                        <td style="text-align: right;">12066</td>
                        <td style="text-align: right;">3500</td>
                        <td style="text-align: right;">5444</td>
                        <td style="text-align: right;">690</td>
                        <td style="text-align: right;">58</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">2</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">桃園市</td>
                        <td style="text-align: center;">2018/10</td>
                        <td style="text-align: right;">27628.0</td>
                        <td style="text-align: right;">1426.7</td>
                        <td style="text-align: right;">2359.1</td>
                        <td style="text-align: right;">81867</td>
                        <td style="text-align: right;">13236</td>
                        <td style="text-align: right;">13811</td>
                        <td style="text-align: right;">51167</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">1</td>
                        <td style="text-align: right;">4</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">4</td>
                        <td style="text-align: right;">1</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr class="alt">
                        <td style="text-align: center;">新竹市</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">1700.0</td>
                        <td style="text-align: right;">2.7</td>
                        <td style="text-align: right;">2.7</td>
                        <td style="text-align: right;">9500</td>
                        <td style="text-align: right;">31</td>
                        <td style="text-align: right;">31</td>
                        <td style="text-align: right;">3120</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">100</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">5</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">新竹縣</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">8000.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">10000</td>
                        <td style="text-align: right;">7</td>
                        <td style="text-align: right;">7</td>
                        <td style="text-align: right;">2500</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">40</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">1</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr class="alt">
                        <td style="text-align: center;">苗栗縣</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">5120.0</td>
                        <td style="text-align: right;">858.0</td>
                        <td style="text-align: right;">233.0</td>
                        <td style="text-align: right;">16000</td>
                        <td style="text-align: right;">3955</td>
                        <td style="text-align: right;">2041</td>
                        <td style="text-align: right;">15000</td>
                        <td style="text-align: right;">200</td>
                        <td style="text-align: right;">200</td>
                        <td style="text-align: right;">15</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">2</td>
                        <td style="text-align: right;">4</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">臺中市</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">9106.0</td>
                        <td style="text-align: right;">5029.7</td>
                        <td style="text-align: right;">5029.7</td>
                        <td style="text-align: right;">113823</td>
                        <td style="text-align: right;">21190</td>
                        <td style="text-align: right;">21154</td>
                        <td style="text-align: right;">284558</td>
                        <td style="text-align: right;">1130</td>
                        <td style="text-align: right;">1130</td>
                        <td style="text-align: right;">128</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">10</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr class="alt">
                        <td style="text-align: center;">彰化縣</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">1327.0</td>
                        <td style="text-align: right;">549.4</td>
                        <td style="text-align: right;">549.4</td>
                        <td style="text-align: right;">16000</td>
                        <td style="text-align: right;">149</td>
                        <td style="text-align: right;">149</td>
                        <td style="text-align: right;">16000</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">10</td>
                        <td style="text-align: right;">2</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">2</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">南投縣</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">3400.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">4400</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">3055</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">40</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">6</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr class="alt">
                        <td style="text-align: center;">雲林縣</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">800.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">17600</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">2000</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">12</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">11</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">嘉義市</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">3200.0</td>
                        <td style="text-align: right;">800.4</td>
                        <td style="text-align: right;">511.7</td>
                        <td style="text-align: right;">10000</td>
                        <td style="text-align: right;">6789</td>
                        <td style="text-align: right;">1129</td>
                        <td style="text-align: right;">1500</td>
                        <td style="text-align: right;">207</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">48</td>
                        <td style="text-align: right;">1</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">3</td>
                        <td style="text-align: right;">2</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr class="alt">
                        <td style="text-align: center;">嘉義縣</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">2000.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">15934</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">20</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;"></td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">臺南市</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">27540.0</td>
                        <td style="text-align: right;">2033.8</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">61560</td>
                        <td style="text-align: right;">4827</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">15623</td>
                        <td style="text-align: right;">112</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">20</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">10</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr class="alt">
                        <td style="text-align: center;">高雄市</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">16972.2</td>
                        <td style="text-align: right;">593.5</td>
                        <td style="text-align: right;">492.0</td>
                        <td style="text-align: right;">56000</td>
                        <td style="text-align: right;">3130</td>
                        <td style="text-align: right;">3085</td>
                        <td style="text-align: right;">56000</td>
                        <td style="text-align: right;">3619</td>
                        <td style="text-align: right;">3619</td>
                        <td style="text-align: right;">20</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">5</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">屏東縣</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">978.0</td>
                        <td style="text-align: right;">25.2</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">8000</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">9600</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">10</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">1</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr class="alt">
                        <td style="text-align: center;">臺東縣</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">3200.3</td>
                        <td style="text-align: right;">107.3</td>
                        <td style="text-align: right;">107.3</td>
                        <td style="text-align: right;">9882</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">2000</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;"></td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;"></td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">花蓮縣</td>
                        <td style="text-align: center;">2018/10</td>
                        <td style="text-align: right;">6000.0</td>
                        <td style="text-align: right;">4151.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">11000</td>
                        <td style="text-align: right;">101</td>
                        <td style="text-align: right;">81</td>
                        <td style="text-align: right;">3000</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">10</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">1</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr class="alt">
                        <td style="text-align: center;">宜蘭縣</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">10132.0</td>
                        <td style="text-align: right;">680.0</td>
                        <td style="text-align: right;">672.0</td>
                        <td style="text-align: right;">38550</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">1728</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">5</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">1</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">澎湖縣</td>
                        <td style="text-align: center;">2018/09</td>
                        <td style="text-align: right;">1250.0</td>
                        <td style="text-align: right;">25.6</td>
                        <td style="text-align: right;">25.6</td>
                        <td style="text-align: right;">4950</td>
                        <td style="text-align: right;">38</td>
                        <td style="text-align: right;">38</td>
                        <td style="text-align: right;">100</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr class="alt">
                        <td style="text-align: center;">金門縣</td>
                        <td style="text-align: center;">2018/10</td>
                        <td style="text-align: right;">1752.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">511</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">200</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">2</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">0.0</td>
                    </tr>
                    <tr>
                        <td style="text-align: center;">合計</td>
                        <td></td>
                        <td style="text-align: right;">253735.5</td>
                        <td style="text-align: right;">39099.0</td>
                        <td style="text-align: right;">14416.2</td>
                        <td style="text-align: right;">843102.0</td>
                        <td style="text-align: right;">100522.0</td>
                        <td style="text-align: right;">55772.0</td>
                        <td style="text-align: right;">661988.0</td>
                        <td style="text-align: right;">25225.0</td>
                        <td style="text-align: right;">7637.0</td>
                        <td style="text-align: right;">746</td>
                        <td style="text-align: right;">27</td>
                        <td style="text-align: right;">0</td>
                        <td style="text-align: right;">98</td>
                        <td style="text-align: right;">7</td>
                        <td style="text-align: right;">0</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
</asp:Content>

