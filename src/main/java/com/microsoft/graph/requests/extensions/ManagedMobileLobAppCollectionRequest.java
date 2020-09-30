// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.DeviceAppManagement;
import com.microsoft.graph.models.extensions.ManagedMobileLobApp;
import java.util.Arrays;
import java.util.EnumSet;

import com.microsoft.graph.options.QueryOption;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseCollectionRequest;
import com.microsoft.graph.concurrency.IExecutors;
import com.microsoft.graph.requests.extensions.IManagedMobileLobAppCollectionPage;
import com.microsoft.graph.requests.extensions.ManagedMobileLobAppCollectionResponse;
import com.microsoft.graph.requests.extensions.IManagedMobileLobAppCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.IManagedMobileLobAppCollectionRequest;
import com.microsoft.graph.requests.extensions.ManagedMobileLobAppCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Managed Mobile Lob App Collection Request.
 */
public class ManagedMobileLobAppCollectionRequest extends BaseCollectionRequest<ManagedMobileLobAppCollectionResponse, IManagedMobileLobAppCollectionPage> implements IManagedMobileLobAppCollectionRequest {

    /**
     * The request builder for this collection of ManagedMobileLobApp
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public ManagedMobileLobAppCollectionRequest(final String requestUrl, IBaseClient client, final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, ManagedMobileLobAppCollectionResponse.class, IManagedMobileLobAppCollectionPage.class);
    }

    public void get(final ICallback<? super IManagedMobileLobAppCollectionPage> callback) {
        final IExecutors executors = getBaseRequest().getClient().getExecutors();
        executors.performOnBackground(new Runnable() {
           @Override
           public void run() {
                try {
                    executors.performOnForeground(get(), callback);
                } catch (final ClientException e) {
                    executors.performOnForeground(e, callback);
                }
           }
        });
    }

    public IManagedMobileLobAppCollectionPage get() throws ClientException {
        final ManagedMobileLobAppCollectionResponse response = send();
        return buildFromResponse(response);
    }

    public void post(final ManagedMobileLobApp newManagedMobileLobApp, final ICallback<? super ManagedMobileLobApp> callback) {
        final String requestUrl = getBaseRequest().getRequestUrl().toString();
        new ManagedMobileLobAppRequestBuilder(requestUrl, getBaseRequest().getClient(), /* Options */ null)
            .buildRequest(getBaseRequest().getHeaders())
            .post(newManagedMobileLobApp, callback);
    }

    public ManagedMobileLobApp post(final ManagedMobileLobApp newManagedMobileLobApp) throws ClientException {
        final String requestUrl = getBaseRequest().getRequestUrl().toString();
        return new ManagedMobileLobAppRequestBuilder(requestUrl, getBaseRequest().getClient(), /* Options */ null)
            .buildRequest(getBaseRequest().getHeaders())
            .post(newManagedMobileLobApp);
    }

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
    public IManagedMobileLobAppCollectionRequest expand(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$expand", value));
        return (ManagedMobileLobAppCollectionRequest)this;
    }

    /**
     * Sets the filter clause for the request
     *
     * @param value the filter clause
     * @return the updated request
     */
    public IManagedMobileLobAppCollectionRequest filter(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$filter", value));
        return (ManagedMobileLobAppCollectionRequest)this;
    }

    /**
     * Sets the order by clause for the request
     *
     * @param value the order by clause
     * @return the updated request
     */
    public IManagedMobileLobAppCollectionRequest orderBy(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$orderby", value));
        return (ManagedMobileLobAppCollectionRequest)this;
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
    public IManagedMobileLobAppCollectionRequest select(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$select", value));
        return (ManagedMobileLobAppCollectionRequest)this;
    }

    /**
     * Sets the top value for the request
     *
     * @param value the max number of items to return
     * @return the updated request
     */
    public IManagedMobileLobAppCollectionRequest top(final int value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$top", value + ""));
        return (ManagedMobileLobAppCollectionRequest)this;
    }

    /**
     * Sets the skip value for the request
     *
     * @param value of the number of items to skip
     * @return the updated request
     */
    public IManagedMobileLobAppCollectionRequest skip(final int value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$skip", value + ""));
        return (ManagedMobileLobAppCollectionRequest)this;
    }


    /**
     * Add Skip token for pagination
     * @param skipToken - Token for pagination
     * @return the updated request
     */
    public IManagedMobileLobAppCollectionRequest skipToken(final String skipToken) {
    	addQueryOption(new QueryOption("$skiptoken", skipToken));
        return (IManagedMobileLobAppCollectionRequest)this;
    }
    public IManagedMobileLobAppCollectionPage buildFromResponse(final ManagedMobileLobAppCollectionResponse response) {
        final IManagedMobileLobAppCollectionRequestBuilder builder;
        if (response.nextLink != null) {
            builder = new ManagedMobileLobAppCollectionRequestBuilder(response.nextLink, getBaseRequest().getClient(), /* options */ null);
        } else {
            builder = null;
        }
        final ManagedMobileLobAppCollectionPage page = new ManagedMobileLobAppCollectionPage(response, builder);
        page.setRawObject(response.getSerializer(), response.getRawObject());
        return page;
    }
}
