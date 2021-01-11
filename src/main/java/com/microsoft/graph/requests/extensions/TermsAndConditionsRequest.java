// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.TermsAndConditions;
import com.microsoft.graph.requests.extensions.TermsAndConditionsAcceptanceStatusCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.TermsAndConditionsAcceptanceStatusRequestBuilder;
import com.microsoft.graph.requests.extensions.TermsAndConditionsAssignmentCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.TermsAndConditionsAssignmentRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Terms And Conditions Request.
 */
public class TermsAndConditionsRequest extends BaseRequest<TermsAndConditions> {
	
    /**
     * The request for the TermsAndConditions
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public TermsAndConditionsRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, TermsAndConditions.class);
    }

    /**
     * Gets the TermsAndConditions from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super TermsAndConditions> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the TermsAndConditions from the service
     *
     * @return the TermsAndConditions from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public TermsAndConditions get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super TermsAndConditions> callback) {
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
     * Patches this TermsAndConditions with a source
     *
     * @param sourceTermsAndConditions the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final TermsAndConditions sourceTermsAndConditions, @Nonnull final ICallback<? super TermsAndConditions> callback) {
        send(HttpMethod.PATCH, callback, sourceTermsAndConditions);
    }

    /**
     * Patches this TermsAndConditions with a source
     *
     * @param sourceTermsAndConditions the source object with updates
     * @return the updated TermsAndConditions
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public TermsAndConditions patch(@Nonnull final TermsAndConditions sourceTermsAndConditions) throws ClientException {
        return send(HttpMethod.PATCH, sourceTermsAndConditions);
    }

    /**
     * Creates a TermsAndConditions with a new object
     *
     * @param newTermsAndConditions the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final TermsAndConditions newTermsAndConditions, @Nonnull final ICallback<? super TermsAndConditions> callback) {
        send(HttpMethod.POST, callback, newTermsAndConditions);
    }

    /**
     * Creates a TermsAndConditions with a new object
     *
     * @param newTermsAndConditions the new object to create
     * @return the created TermsAndConditions
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public TermsAndConditions post(@Nonnull final TermsAndConditions newTermsAndConditions) throws ClientException {
        return send(HttpMethod.POST, newTermsAndConditions);
    }

    /**
     * Creates a TermsAndConditions with a new object
     *
     * @param newTermsAndConditions the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final TermsAndConditions newTermsAndConditions, @Nonnull final ICallback<? super TermsAndConditions> callback) {
        send(HttpMethod.PUT, callback, newTermsAndConditions);
    }

    /**
     * Creates a TermsAndConditions with a new object
     *
     * @param newTermsAndConditions the object to create/update
     * @return the created TermsAndConditions
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public TermsAndConditions put(@Nonnull final TermsAndConditions newTermsAndConditions) throws ClientException {
        return send(HttpMethod.PUT, newTermsAndConditions);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public TermsAndConditionsRequest select(@Nonnull final String value) {
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
     public TermsAndConditionsRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

