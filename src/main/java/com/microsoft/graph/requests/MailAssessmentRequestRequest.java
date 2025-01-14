// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.MailAssessmentRequest;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Mail Assessment Request Request.
 */
public class MailAssessmentRequestRequest extends BaseRequest<MailAssessmentRequest> {
	
    /**
     * The request for the MailAssessmentRequest
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public MailAssessmentRequestRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, MailAssessmentRequest.class);
    }

    /**
     * Gets the MailAssessmentRequest from the service
     *
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<MailAssessmentRequest> getAsync() {
        return sendAsync(HttpMethod.GET, null);
    }

    /**
     * Gets the MailAssessmentRequest from the service
     *
     * @return the MailAssessmentRequest from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public MailAssessmentRequest get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @return a future with the deletion result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<MailAssessmentRequest> deleteAsync() {
        return sendAsync(HttpMethod.DELETE, null);
    }

    /**
     * Delete this item from the service
     * @return the resulting response if the service returns anything on deletion
     *
     * @throws ClientException if there was an exception during the delete operation
     */
    @Nullable
    public MailAssessmentRequest delete() throws ClientException {
        return send(HttpMethod.DELETE, null);
    }

    /**
     * Patches this MailAssessmentRequest with a source
     *
     * @param sourceMailAssessmentRequest the source object with updates
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<MailAssessmentRequest> patchAsync(@Nonnull final MailAssessmentRequest sourceMailAssessmentRequest) {
        return sendAsync(HttpMethod.PATCH, sourceMailAssessmentRequest);
    }

    /**
     * Patches this MailAssessmentRequest with a source
     *
     * @param sourceMailAssessmentRequest the source object with updates
     * @return the updated MailAssessmentRequest
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public MailAssessmentRequest patch(@Nonnull final MailAssessmentRequest sourceMailAssessmentRequest) throws ClientException {
        return send(HttpMethod.PATCH, sourceMailAssessmentRequest);
    }

    /**
     * Creates a MailAssessmentRequest with a new object
     *
     * @param newMailAssessmentRequest the new object to create
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<MailAssessmentRequest> postAsync(@Nonnull final MailAssessmentRequest newMailAssessmentRequest) {
        return sendAsync(HttpMethod.POST, newMailAssessmentRequest);
    }

    /**
     * Creates a MailAssessmentRequest with a new object
     *
     * @param newMailAssessmentRequest the new object to create
     * @return the created MailAssessmentRequest
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public MailAssessmentRequest post(@Nonnull final MailAssessmentRequest newMailAssessmentRequest) throws ClientException {
        return send(HttpMethod.POST, newMailAssessmentRequest);
    }

    /**
     * Creates a MailAssessmentRequest with a new object
     *
     * @param newMailAssessmentRequest the object to create/update
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<MailAssessmentRequest> putAsync(@Nonnull final MailAssessmentRequest newMailAssessmentRequest) {
        return sendAsync(HttpMethod.PUT, newMailAssessmentRequest);
    }

    /**
     * Creates a MailAssessmentRequest with a new object
     *
     * @param newMailAssessmentRequest the object to create/update
     * @return the created MailAssessmentRequest
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public MailAssessmentRequest put(@Nonnull final MailAssessmentRequest newMailAssessmentRequest) throws ClientException {
        return send(HttpMethod.PUT, newMailAssessmentRequest);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public MailAssessmentRequestRequest select(@Nonnull final String value) {
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
     public MailAssessmentRequestRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

