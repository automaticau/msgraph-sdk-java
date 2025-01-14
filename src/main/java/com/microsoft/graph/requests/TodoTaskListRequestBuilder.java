// Template Source: BaseEntityRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.TodoTaskList;
import com.microsoft.graph.requests.ExtensionCollectionRequestBuilder;
import com.microsoft.graph.requests.ExtensionRequestBuilder;
import com.microsoft.graph.requests.TodoTaskCollectionRequestBuilder;
import com.microsoft.graph.requests.TodoTaskRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequestBuilder;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Todo Task List Request Builder.
 */
public class TodoTaskListRequestBuilder extends BaseRequestBuilder<TodoTaskList> {

    /**
     * The request builder for the TodoTaskList
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public TodoTaskListRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }

    /**
     * Creates the request
     *
     * @param requestOptions the options for this request
     * @return the TodoTaskListRequest instance
     */
    @Nonnull
    public TodoTaskListRequest buildRequest(@Nullable final com.microsoft.graph.options.Option... requestOptions) {
        return buildRequest(getOptions(requestOptions));
    }

    /**
     * Creates the request with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for this request
     * @return the TodoTaskListRequest instance
     */
    @Nonnull
    public TodoTaskListRequest buildRequest(@Nonnull final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        return new com.microsoft.graph.requests.TodoTaskListRequest(getRequestUrl(), getClient(), requestOptions);
    }


    /**
     *  Gets a request builder for the Extension collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public ExtensionCollectionRequestBuilder extensions() {
        return new ExtensionCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("extensions"), getClient(), null);
    }

    /**
     * Gets a request builder for the Extension item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public ExtensionRequestBuilder extensions(@Nonnull final String id) {
        return new ExtensionRequestBuilder(getRequestUrlWithAdditionalSegment("extensions") + "/" + id, getClient(), null);
    }
    /**
     *  Gets a request builder for the TodoTask collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public TodoTaskCollectionRequestBuilder tasks() {
        return new TodoTaskCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("tasks"), getClient(), null);
    }

    /**
     * Gets a request builder for the TodoTask item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public TodoTaskRequestBuilder tasks(@Nonnull final String id) {
        return new TodoTaskRequestBuilder(getRequestUrlWithAdditionalSegment("tasks") + "/" + id, getClient(), null);
    }
}
