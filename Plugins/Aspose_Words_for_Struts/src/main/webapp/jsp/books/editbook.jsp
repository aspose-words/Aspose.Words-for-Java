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
	<b>Edit Book</b>
	<html:form>
		<table style="background-color: red; border: 1px solid black">
			<tr>
				<td>Book Id</td>
				<td><html:text property="bookId" disabled="true" /></td>
			</tr>
			<tr>
				<td>Book Name</td>
				<td><html:text property="bookName" /></td>
			</tr>
			<tr>
				<td>Author Name</td>
				<td><html:text property="authorName" /></td>
			</tr>
			<tr>
				<td>Book Cost</td>
				<td><html:text property="bookCost" /></td>
			</tr>
		</table>
		</p>
		<p>
		<table>
			<tr>
				<td><input type="submit" name="actionMethod" value="UpdateBook" /></td>
			</tr>
		</table>
	</html:form>
	</p>
</body>
</html>