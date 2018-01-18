// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

// **NOTE** This file was generated by a tool and any changes will be overwritten.
package com.microsoft.graph.requests.generated;

import com.microsoft.graph.concurrency.*;
import com.microsoft.graph.core.*;
import com.microsoft.graph.models.extensions.*;
import com.microsoft.graph.models.generated.*;
import com.microsoft.graph.http.*;
import com.microsoft.graph.requests.extensions.*;
import com.microsoft.graph.requests.generated.*;
import com.microsoft.graph.options.*;
import com.microsoft.graph.serializer.*;

import java.util.Arrays;
import java.util.EnumSet;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Base Onenote Section Collection Request.
 */
public class BaseOnenoteSectionCollectionRequest extends BaseCollectionRequest<BaseOnenoteSectionCollectionResponse, IOnenoteSectionCollectionPage> implements IBaseOnenoteSectionCollectionRequest {

    /**
     * The request builder for this collection of OnenoteSection
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     */
    public BaseOnenoteSectionCollectionRequest(final String requestUrl, IBaseClient client, final java.util.List<? extends Option> requestOptions) {
        super(requestUrl, client, requestOptions, BaseOnenoteSectionCollectionResponse.class, IOnenoteSectionCollectionPage.class);
    }

    public void get(final ICallback<IOnenoteSectionCollectionPage> callback) {
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

    public IOnenoteSectionCollectionPage get() throws ClientException {
        final BaseOnenoteSectionCollectionResponse response = send();
        return buildFromResponse(response);
    }

    public void post(final OnenoteSection newOnenoteSection, final ICallback<OnenoteSection> callback) {
        final String requestUrl = getBaseRequest().getRequestUrl().toString();
        new OnenoteSectionRequestBuilder(requestUrl, getBaseRequest().getClient(), /* Options */ null)
            .buildRequest(getBaseRequest().getOptions())
            .post(newOnenoteSection, callback);
    }

    public OnenoteSection post(final OnenoteSection newOnenoteSection) throws ClientException {
        final String requestUrl = getBaseRequest().getRequestUrl().toString();
        return new OnenoteSectionRequestBuilder(requestUrl, getBaseRequest().getClient(), /* Options */ null)
            .buildRequest(getBaseRequest().getOptions())
            .post(newOnenoteSection);
    }

    /**
     * Sets the expand clause for the request
     *
     * @param value The expand clause
     * @return The updated request
     */
    public IOnenoteSectionCollectionRequest expand(final String value) {
        addQueryOption(new QueryOption("$expand", value));
        return (OnenoteSectionCollectionRequest)this;
    }

    /**
     * Sets the select clause for the request
     *
     * @param value The select clause
     * @return The updated request
     */
    public IOnenoteSectionCollectionRequest select(final String value) {
        addQueryOption(new QueryOption("$select", value));
        return (OnenoteSectionCollectionRequest)this;
    }

    /**
     * Sets the top value for the request
     *
     * @param value The max number of items to return
     * @return The updated request
     */
    public IOnenoteSectionCollectionRequest top(final int value) {
        addQueryOption(new QueryOption("$top", value + ""));
        return (OnenoteSectionCollectionRequest)this;
    }

    public IOnenoteSectionCollectionPage buildFromResponse(final BaseOnenoteSectionCollectionResponse response) {
        final IOnenoteSectionCollectionRequestBuilder builder;
        if (response.nextLink != null) {
            builder = new OnenoteSectionCollectionRequestBuilder(response.nextLink, getBaseRequest().getClient(), /* options */ null);
        } else {
            builder = null;
        }
        final OnenoteSectionCollectionPage page = new OnenoteSectionCollectionPage(response, builder);
        page.setRawObject(response.getSerializer(), response.getRawObject());
        return page;
    }
}
