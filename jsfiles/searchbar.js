﻿document.write ('<table>');
document.write ('<form method="post" name="SearchForm" action="/NBArticle/NBArticle/search.asp?action=query" target="_blank">');
document.write ('<tr>');
document.write ('<td align="center"><span style="color: #000000;">站内搜索：</span></td>');
document.write ('<td align="center">&nbsp;');
document.write ('<select name="field">');
document.write ('<option value="0">标题</option>');
document.write ('<option value="1">关键字</option>');
document.write ('<option value="2">作者</option>');
document.write ('<option value="3">摘要</option>');
document.write ('</select>&nbsp;');
document.write ('<select name="column">');
document.write ('<option value="0">--栏 目--</option>');
document.write ('├');
document.write ('');
document.write ('</select>&nbsp;<input name="keyword" type="text" value="关键字" onfocus="this.select();" size="20" maxlength="50">&nbsp;<input name="Submit" type="submit" value="搜索"></td>');
document.write ('<td align="center">&nbsp;<a href="/NBArticle/NBArticle/search.asp">高级搜索</a></td>');
document.write ('</tr>');
document.write ('</form>');
document.write ('</table>');
