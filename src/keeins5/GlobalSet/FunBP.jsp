	</table>
<%if(bpfun==0) %><input type="submit" value="確定">　<input type="reset" value="清除">
<%else if(bpfun==2) %><input type="submit" value="確定">
<%else if(bpfun==3) %><input type="submit" value="確定">　<input type="reset" value="清除">　<input type="reset" name="s2" value="取消">
<%else if(bpfun==4) {%><input type="submit" value="列印" name="Print"><%}%>
</form>
</center>
<%@include file = "../GlobalSet/BPanel.jsp"%>
</body>
</html>
