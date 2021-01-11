// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.models.extensions.ChatMessageHostedContent;
import com.microsoft.graph.requests.extensions.ChatMessageHostedContentCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.extensions.ChatMessageHostedContentCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Chat Message Hosted Content Collection Page.
 */
public class ChatMessageHostedContentCollectionPage extends BaseCollectionPage<ChatMessageHostedContent, ChatMessageHostedContentCollectionRequestBuilder> {

    /**
     * A collection page for ChatMessageHostedContent
     *
     * @param response the serialized ChatMessageHostedContentCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public ChatMessageHostedContentCollectionPage(@Nonnull final ChatMessageHostedContentCollectionResponse response, @Nonnull final ChatMessageHostedContentCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for ChatMessageHostedContent
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public ChatMessageHostedContentCollectionPage(@Nonnull final java.util.List<ChatMessageHostedContent> pageContents, @Nullable final ChatMessageHostedContentCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
