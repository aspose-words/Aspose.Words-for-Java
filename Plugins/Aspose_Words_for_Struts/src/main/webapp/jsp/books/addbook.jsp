<%@taglib uri="http://struts.apache.org/tags-html" prefix="html"%>
<html>
<body>
	<p>
	<h2>
		<img src="images/aspose-struts-logo.jpg"
			alt="Aspose.Words for Java"> Aspose.Words Struts Example -
		Simple Book Store App
	</h2>
	</p>
	<b>Add Book</b>
	<html:form>
		<table style="background-color: red; border: 1px solid black">
			<tr>
				<td>Book Name</td>
				<td><html:text property="bookName" value="" /></td>
			</tr>
			<tr>
				<td>Author Name</td>
				<td><html:text property="authorName" value="" /></td>
			</tr>
			<tr>
				<td>Book Cost</td>
				<td><html:text property="bookCost" value="" /></td>
			</tr>
		</table>
		</p>
		<p>
		<table>
			<tr>
				<td><input type="submit" name="actionMethod" value="SaveBook" /></td>
			</tr>
		</table>
	</html:form>
	</p>
</body>
</html>