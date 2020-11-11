package com.microsoft.graph.functional;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;

import java.io.InputStream;
import java.util.LinkedList;

import com.microsoft.graph.concurrency.ChunkedUploadProvider;
import com.microsoft.graph.concurrency.IProgressCallback;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.extensions.Channel;
import com.microsoft.graph.models.extensions.ChatMessage;
import com.microsoft.graph.models.extensions.ChatMessageAttachment;
import com.microsoft.graph.models.extensions.DriveItem;
import com.microsoft.graph.models.extensions.DriveItemUploadableProperties;
import com.microsoft.graph.models.extensions.ItemBody;
import com.microsoft.graph.models.extensions.UploadSession;
import com.microsoft.graph.models.generated.BodyType;
import com.microsoft.graph.requests.extensions.IGroupCollectionPage;

@Ignore
public class TeamsTests {
	@Test
    public void createMessageWithLargeAttachements() throws Exception {
    	TestBase testBase = new TestBase();
		InputStream uploadFile = OneDriveTests.class.getClassLoader().getResourceAsStream("largefile10M.blob");
		long fileSize = (long) uploadFile.available();

		IGroupCollectionPage teams = testBase.graphClient.groups().buildRequest().top(1).filter("mail eq 'TBR@baywet.onmicrosoft.com'").get();
		String teamId = teams.getCurrentPage().get(0).id;
		
		UploadSession uploadSession = testBase
				.graphClient
				.groups(teamId)
				.drive()
				.root()
				.itemWithPath("largefile10M.blob")
				.createUploadSession(new DriveItemUploadableProperties())
				.buildRequest()
				.post();

		IProgressCallback<DriveItem> callback = new IProgressCallback<DriveItem> () {
			@Override
			public void progress(final long current, final long max) {
				//Check progress
			}
			@Override
			public void success(final DriveItem result) {
				//Handle the successful response
				String finishedItemId = result.id;
				Assert.assertNotNull(finishedItemId);
				ChatMessage message = new ChatMessage();
				message.body = new ItemBody();
				message.body.contentType = BodyType.HTML;
				message.body.content = "Here's the latest budget. <attachment id=\"153fa47d-18c9-4179-be08-9879815a9f90\"></attachment>";
				message.attachments = new LinkedList<ChatMessageAttachment>();
				ChatMessageAttachment atta = new ChatMessageAttachment();
				atta.id = "153fa47d-18c9-4179-be08-9879815a9f90";
				atta.contentType = "reference";
				atta.name = "largefile10M.blob";
				atta.contentUrl = result.webUrl;
				message.attachments.add(atta);

				Channel primaryChannel = testBase.graphClient.teams(teamId).primaryChannel().buildRequest().get(); // the API doesn't support creating message on that endpoint directly

				testBase.graphClient.teams(teamId).channels(primaryChannel.id).messages().buildRequest().post(message);
			}
			
			@Override
			public void failure(final ClientException ex) {
				//Handle the failed upload
				Assert.fail("Upload session failed");
			}
		};
		ChunkedUploadProvider<DriveItem> chunkedUploadProvider = new ChunkedUploadProvider<DriveItem>(
				uploadSession, 
				testBase.graphClient, 
				uploadFile, 
				fileSize, 
				DriveItem.class);
		
		chunkedUploadProvider.upload(callback);
    }
}
