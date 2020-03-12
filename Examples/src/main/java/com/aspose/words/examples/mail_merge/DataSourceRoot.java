package com.aspose.words.examples.mail_merge;

import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.IMailMergeDataSourceRoot;
import com.aspose.words.ref.Ref;

public class DataSourceRoot implements IMailMergeDataSourceRoot {
	@Override
	public IMailMergeDataSource getDataSource(String s) throws Exception {
		return new DataSource();
	}

	private class DataSource implements IMailMergeDataSource {

		boolean next = true;

		@Override
		public String getTableName() throws Exception {
			return "example";
		}

		@Override
		public boolean moveNext() throws Exception {
			boolean result = next;
			next = false;
			return result;
		}

		@Override
		public boolean getValue(String s, Ref<Object> ref) throws Exception {
			return false;
		}

		@Override
		public IMailMergeDataSource getChildDataSource(String s) throws Exception {
			return null;
		}
	}
}