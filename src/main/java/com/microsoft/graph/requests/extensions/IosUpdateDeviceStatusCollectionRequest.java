// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.DeviceManagement;
import com.microsoft.graph.models.extensions.IosUpdateDeviceStatus;
import java.util.Arrays;
import java.util.EnumSet;

import com.microsoft.graph.options.QueryOption;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseCollectionRequest;
import com.microsoft.graph.concurrency.IExecutors;
import com.microsoft.graph.requests.extensions.IosUpdateDeviceStatusCollectionResponse;
import com.microsoft.graph.requests.extensions.IosUpdateDeviceStatusCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.IosUpdateDeviceStatusCollectionRequest;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Ios Update Device Status Collection Request.
 */
public class IosUpdateDeviceStatusCollectionRequest extends BaseCollectionRequest<IosUpdateDeviceStatus, IosUpdateDeviceStatusCollectionResponse, IosUpdateDeviceStatusCollectionPage> {

    /**
     * The request builder for this collection of IosUpdateDeviceStatus
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public IosUpdateDeviceStatusCollectionRequest(final String requestUrl, IBaseClient client, final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, IosUpdateDeviceStatusCollectionResponse.class, IosUpdateDeviceStatusCollectionPage.class, IosUpdateDeviceStatusCollectionRequestBuilder.class);
    }

    public void post(final IosUpdateDeviceStatus newIosUpdateDeviceStatus, final ICallback<? super IosUpdateDeviceStatus> callback) {
        final String requestUrl = getBaseRequest().getRequestUrl().toString();
        new IosUpdateDeviceStatusRequestBuilder(requestUrl, getBaseRequest().getClient(), /* Options */ null)
            .buildRequest(getBaseRequest().getHeaders())
            .post(newIosUpdateDeviceStatus, callback);
    }

    public IosUpdateDeviceStatus post(final IosUpdateDeviceStatus newIosUpdateDeviceStatus) throws ClientException {
        final String requestUrl = getBaseRequest().getRequestUrl().toString();
        return new IosUpdateDeviceStatusRequestBuilder(requestUrl, getBaseRequest().getClient(), /* Options */ null)
            .buildRequest(getBaseRequest().getHeaders())
            .post(newIosUpdateDeviceStatus);
    }

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
    public IosUpdateDeviceStatusCollectionRequest expand(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$expand", value));
        return this;
    }

    /**
     * Sets the filter clause for the request
     *
     * @param value the filter clause
     * @return the updated request
     */
    public IosUpdateDeviceStatusCollectionRequest filter(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$filter", value));
        return this;
    }

    /**
     * Sets the order by clause for the request
     *
     * @param value the order by clause
     * @return the updated request
     */
    public IosUpdateDeviceStatusCollectionRequest orderBy(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$orderby", value));
        return this;
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
    public IosUpdateDeviceStatusCollectionRequest select(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$select", value));
        return this;
    }

    /**
     * Sets the top value for the request
     *
     * @param value the max number of items to return
     * @return the updated request
     */
    public IosUpdateDeviceStatusCollectionRequest top(final int value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$top", value + ""));
        return this;
    }

    /**
     * Sets the skip value for the request
     *
     * @param value of the number of items to skip
     * @return the updated request
     */
    public IosUpdateDeviceStatusCollectionRequest skip(final int value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$skip", value + ""));
        return this;
    }


    /**
     * Add Skip token for pagination
     * @param skipToken - Token for pagination
     * @return the updated request
     */
    public IosUpdateDeviceStatusCollectionRequest skipToken(final String skipToken) {
    	addQueryOption(new QueryOption("$skiptoken", skipToken));
        return this;
    }
}
