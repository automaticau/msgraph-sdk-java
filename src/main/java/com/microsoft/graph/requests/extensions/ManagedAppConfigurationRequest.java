// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.ManagedAppConfiguration;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Managed App Configuration Request.
 */
public class ManagedAppConfigurationRequest extends BaseRequest<ManagedAppConfiguration> {
	
    /**
     * The request for the ManagedAppConfiguration
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     * @param responseClass  the class of the response
     */
    public ManagedAppConfigurationRequest(@Nonnull final String requestUrl,
            @Nonnull final IBaseClient client,
            @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions,
            @Nonnull final Class<? extends ManagedAppConfiguration> responseClass) {
        super(requestUrl, client, requestOptions, responseClass);
    }

    /**
     * The request for the ManagedAppConfiguration
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public ManagedAppConfigurationRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, ManagedAppConfiguration.class);
    }

    /**
     * Gets the ManagedAppConfiguration from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super ManagedAppConfiguration> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the ManagedAppConfiguration from the service
     *
     * @return the ManagedAppConfiguration from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public ManagedAppConfiguration get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super ManagedAppConfiguration> callback) {
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
     * Patches this ManagedAppConfiguration with a source
     *
     * @param sourceManagedAppConfiguration the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final ManagedAppConfiguration sourceManagedAppConfiguration, @Nonnull final ICallback<? super ManagedAppConfiguration> callback) {
        send(HttpMethod.PATCH, callback, sourceManagedAppConfiguration);
    }

    /**
     * Patches this ManagedAppConfiguration with a source
     *
     * @param sourceManagedAppConfiguration the source object with updates
     * @return the updated ManagedAppConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public ManagedAppConfiguration patch(@Nonnull final ManagedAppConfiguration sourceManagedAppConfiguration) throws ClientException {
        return send(HttpMethod.PATCH, sourceManagedAppConfiguration);
    }

    /**
     * Creates a ManagedAppConfiguration with a new object
     *
     * @param newManagedAppConfiguration the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final ManagedAppConfiguration newManagedAppConfiguration, @Nonnull final ICallback<? super ManagedAppConfiguration> callback) {
        send(HttpMethod.POST, callback, newManagedAppConfiguration);
    }

    /**
     * Creates a ManagedAppConfiguration with a new object
     *
     * @param newManagedAppConfiguration the new object to create
     * @return the created ManagedAppConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public ManagedAppConfiguration post(@Nonnull final ManagedAppConfiguration newManagedAppConfiguration) throws ClientException {
        return send(HttpMethod.POST, newManagedAppConfiguration);
    }

    /**
     * Creates a ManagedAppConfiguration with a new object
     *
     * @param newManagedAppConfiguration the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final ManagedAppConfiguration newManagedAppConfiguration, @Nonnull final ICallback<? super ManagedAppConfiguration> callback) {
        send(HttpMethod.PUT, callback, newManagedAppConfiguration);
    }

    /**
     * Creates a ManagedAppConfiguration with a new object
     *
     * @param newManagedAppConfiguration the object to create/update
     * @return the created ManagedAppConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public ManagedAppConfiguration put(@Nonnull final ManagedAppConfiguration newManagedAppConfiguration) throws ClientException {
        return send(HttpMethod.PUT, newManagedAppConfiguration);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public ManagedAppConfigurationRequest select(@Nonnull final String value) {
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
     public ManagedAppConfigurationRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

