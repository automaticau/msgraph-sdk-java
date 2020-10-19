// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.InferenceClassification;
import com.microsoft.graph.models.extensions.InferenceClassificationOverride;
import java.util.Arrays;
import java.util.EnumSet;

import com.microsoft.graph.options.QueryOption;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseCollectionRequest;
import com.microsoft.graph.concurrency.IExecutors;
import com.microsoft.graph.requests.extensions.InferenceClassificationOverrideCollectionResponse;
import com.microsoft.graph.requests.extensions.InferenceClassificationOverrideCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.InferenceClassificationOverrideCollectionRequest;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Inference Classification Override Collection Request.
 */
public class InferenceClassificationOverrideCollectionRequest extends BaseCollectionRequest<InferenceClassificationOverride, InferenceClassificationOverrideCollectionResponse, InferenceClassificationOverrideCollectionPage> {

    /**
     * The request builder for this collection of InferenceClassificationOverride
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public InferenceClassificationOverrideCollectionRequest(final String requestUrl, IBaseClient client, final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, InferenceClassificationOverrideCollectionResponse.class, InferenceClassificationOverrideCollectionPage.class, InferenceClassificationOverrideCollectionRequestBuilder.class);
    }

    public void post(final InferenceClassificationOverride newInferenceClassificationOverride, final ICallback<? super InferenceClassificationOverride> callback) {
        final String requestUrl = getBaseRequest().getRequestUrl().toString();
        new InferenceClassificationOverrideRequestBuilder(requestUrl, getBaseRequest().getClient(), /* Options */ null)
            .buildRequest(getBaseRequest().getHeaders())
            .post(newInferenceClassificationOverride, callback);
    }

    public InferenceClassificationOverride post(final InferenceClassificationOverride newInferenceClassificationOverride) throws ClientException {
        final String requestUrl = getBaseRequest().getRequestUrl().toString();
        return new InferenceClassificationOverrideRequestBuilder(requestUrl, getBaseRequest().getClient(), /* Options */ null)
            .buildRequest(getBaseRequest().getHeaders())
            .post(newInferenceClassificationOverride);
    }

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
    public InferenceClassificationOverrideCollectionRequest expand(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$expand", value));
        return this;
    }

    /**
     * Sets the filter clause for the request
     *
     * @param value the filter clause
     * @return the updated request
     */
    public InferenceClassificationOverrideCollectionRequest filter(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$filter", value));
        return this;
    }

    /**
     * Sets the order by clause for the request
     *
     * @param value the order by clause
     * @return the updated request
     */
    public InferenceClassificationOverrideCollectionRequest orderBy(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$orderby", value));
        return this;
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
    public InferenceClassificationOverrideCollectionRequest select(final String value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$select", value));
        return this;
    }

    /**
     * Sets the top value for the request
     *
     * @param value the max number of items to return
     * @return the updated request
     */
    public InferenceClassificationOverrideCollectionRequest top(final int value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$top", value + ""));
        return this;
    }

    /**
     * Sets the skip value for the request
     *
     * @param value of the number of items to skip
     * @return the updated request
     */
    public InferenceClassificationOverrideCollectionRequest skip(final int value) {
        addQueryOption(new com.microsoft.graph.options.QueryOption("$skip", value + ""));
        return this;
    }


    /**
     * Add Skip token for pagination
     * @param skipToken - Token for pagination
     * @return the updated request
     */
    public InferenceClassificationOverrideCollectionRequest skipToken(final String skipToken) {
    	addQueryOption(new QueryOption("$skiptoken", skipToken));
        return this;
    }
}
