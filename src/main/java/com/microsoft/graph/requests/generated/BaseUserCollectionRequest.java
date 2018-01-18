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
 * The class for the Base User Collection Request.
 */
public class BaseUserCollectionRequest extends BaseCollectionRequest<BaseUserCollectionResponse, IUserCollectionPage> implements IBaseUserCollectionRequest {

    /**
     * The request builder for this collection of User
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     */
    public BaseUserCollectionRequest(final String requestUrl, IBaseClient client, final java.util.List<? extends Option> requestOptions) {
        super(requestUrl, client, requestOptions, BaseUserCollectionResponse.class, IUserCollectionPage.class);
    }

    public void get(final ICallback<IUserCollectionPage> callback) {
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

    public IUserCollectionPage get() throws ClientException {
        final BaseUserCollectionResponse response = send();
        return buildFromResponse(response);
    }

    public void post(final User newUser, final ICallback<User> callback) {
        final String requestUrl = getBaseRequest().getRequestUrl().toString();
        new UserRequestBuilder(requestUrl, getBaseRequest().getClient(), /* Options */ null)
            .buildRequest(getBaseRequest().getOptions())
            .post(newUser, callback);
    }

    public User post(final User newUser) throws ClientException {
        final String requestUrl = getBaseRequest().getRequestUrl().toString();
        return new UserRequestBuilder(requestUrl, getBaseRequest().getClient(), /* Options */ null)
            .buildRequest(getBaseRequest().getOptions())
            .post(newUser);
    }

    /**
     * Sets the expand clause for the request
     *
     * @param value The expand clause
     * @return The updated request
     */
    public IUserCollectionRequest expand(final String value) {
        addQueryOption(new QueryOption("$expand", value));
        return (UserCollectionRequest)this;
    }

    /**
     * Sets the select clause for the request
     *
     * @param value The select clause
     * @return The updated request
     */
    public IUserCollectionRequest select(final String value) {
        addQueryOption(new QueryOption("$select", value));
        return (UserCollectionRequest)this;
    }

    /**
     * Sets the top value for the request
     *
     * @param value The max number of items to return
     * @return The updated request
     */
    public IUserCollectionRequest top(final int value) {
        addQueryOption(new QueryOption("$top", value + ""));
        return (UserCollectionRequest)this;
    }

    public IUserCollectionPage buildFromResponse(final BaseUserCollectionResponse response) {
        final IUserCollectionRequestBuilder builder;
        if (response.nextLink != null) {
            builder = new UserCollectionRequestBuilder(response.nextLink, getBaseRequest().getClient(), /* options */ null);
        } else {
            builder = null;
        }
        final UserCollectionPage page = new UserCollectionPage(response, builder);
        page.setRawObject(response.getSerializer(), response.getRawObject());
        return page;
    }
}
