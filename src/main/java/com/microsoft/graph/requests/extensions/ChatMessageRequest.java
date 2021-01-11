// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.ChatMessage;
import com.microsoft.graph.requests.extensions.ChatMessageHostedContentCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.ChatMessageHostedContentRequestBuilder;
import com.microsoft.graph.requests.extensions.ChatMessageCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.ChatMessageRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Chat Message Request.
 */
public class ChatMessageRequest extends BaseRequest<ChatMessage> {
	
    /**
     * The request for the ChatMessage
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public ChatMessageRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, ChatMessage.class);
    }

    /**
     * Gets the ChatMessage from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super ChatMessage> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the ChatMessage from the service
     *
     * @return the ChatMessage from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public ChatMessage get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super ChatMessage> callback) {
        send(HttpMethod.DELETE, callback, null);
    }

    /**
     * Delete this item from the service
     *
     * @throws ClientException if there was an exception during the delete operation
     */
    public void delete() throws ClientException {
        send(HttpMethod.DELETE, null);
    }

    /**
     * Patches this ChatMessage with a source
     *
     * @param sourceChatMessage the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final ChatMessage sourceChatMessage, @Nonnull final ICallback<? super ChatMessage> callback) {
        send(HttpMethod.PATCH, callback, sourceChatMessage);
    }

    /**
     * Patches this ChatMessage with a source
     *
     * @param sourceChatMessage the source object with updates
     * @return the updated ChatMessage
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public ChatMessage patch(@Nonnull final ChatMessage sourceChatMessage) throws ClientException {
        return send(HttpMethod.PATCH, sourceChatMessage);
    }

    /**
     * Creates a ChatMessage with a new object
     *
     * @param newChatMessage the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final ChatMessage newChatMessage, @Nonnull final ICallback<? super ChatMessage> callback) {
        send(HttpMethod.POST, callback, newChatMessage);
    }

    /**
     * Creates a ChatMessage with a new object
     *
     * @param newChatMessage the new object to create
     * @return the created ChatMessage
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public ChatMessage post(@Nonnull final ChatMessage newChatMessage) throws ClientException {
        return send(HttpMethod.POST, newChatMessage);
    }

    /**
     * Creates a ChatMessage with a new object
     *
     * @param newChatMessage the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final ChatMessage newChatMessage, @Nonnull final ICallback<? super ChatMessage> callback) {
        send(HttpMethod.PUT, callback, newChatMessage);
    }

    /**
     * Creates a ChatMessage with a new object
     *
     * @param newChatMessage the object to create/update
     * @return the created ChatMessage
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public ChatMessage put(@Nonnull final ChatMessage newChatMessage) throws ClientException {
        return send(HttpMethod.PUT, newChatMessage);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public ChatMessageRequest select(@Nonnull final String value) {
         addSelectOption(value);
         return this;
     }

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
     @Nonnull
     public ChatMessageRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

