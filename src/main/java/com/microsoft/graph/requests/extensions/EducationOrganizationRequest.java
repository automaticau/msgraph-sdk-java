// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.EducationOrganization;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Education Organization Request.
 */
public class EducationOrganizationRequest extends BaseRequest<EducationOrganization> {
	
    /**
     * The request for the EducationOrganization
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     * @param responseClass  the class of the response
     */
    public EducationOrganizationRequest(@Nonnull final String requestUrl,
            @Nonnull final IBaseClient client,
            @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions,
            @Nonnull final Class<? extends EducationOrganization> responseClass) {
        super(requestUrl, client, requestOptions, responseClass);
    }

    /**
     * The request for the EducationOrganization
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public EducationOrganizationRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, EducationOrganization.class);
    }

    /**
     * Gets the EducationOrganization from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super EducationOrganization> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the EducationOrganization from the service
     *
     * @return the EducationOrganization from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public EducationOrganization get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super EducationOrganization> callback) {
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
     * Patches this EducationOrganization with a source
     *
     * @param sourceEducationOrganization the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final EducationOrganization sourceEducationOrganization, @Nonnull final ICallback<? super EducationOrganization> callback) {
        send(HttpMethod.PATCH, callback, sourceEducationOrganization);
    }

    /**
     * Patches this EducationOrganization with a source
     *
     * @param sourceEducationOrganization the source object with updates
     * @return the updated EducationOrganization
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public EducationOrganization patch(@Nonnull final EducationOrganization sourceEducationOrganization) throws ClientException {
        return send(HttpMethod.PATCH, sourceEducationOrganization);
    }

    /**
     * Creates a EducationOrganization with a new object
     *
     * @param newEducationOrganization the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final EducationOrganization newEducationOrganization, @Nonnull final ICallback<? super EducationOrganization> callback) {
        send(HttpMethod.POST, callback, newEducationOrganization);
    }

    /**
     * Creates a EducationOrganization with a new object
     *
     * @param newEducationOrganization the new object to create
     * @return the created EducationOrganization
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public EducationOrganization post(@Nonnull final EducationOrganization newEducationOrganization) throws ClientException {
        return send(HttpMethod.POST, newEducationOrganization);
    }

    /**
     * Creates a EducationOrganization with a new object
     *
     * @param newEducationOrganization the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final EducationOrganization newEducationOrganization, @Nonnull final ICallback<? super EducationOrganization> callback) {
        send(HttpMethod.PUT, callback, newEducationOrganization);
    }

    /**
     * Creates a EducationOrganization with a new object
     *
     * @param newEducationOrganization the object to create/update
     * @return the created EducationOrganization
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public EducationOrganization put(@Nonnull final EducationOrganization newEducationOrganization) throws ClientException {
        return send(HttpMethod.PUT, newEducationOrganization);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public EducationOrganizationRequest select(@Nonnull final String value) {
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
     public EducationOrganizationRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

