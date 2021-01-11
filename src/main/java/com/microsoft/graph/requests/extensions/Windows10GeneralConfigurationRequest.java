// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.Windows10GeneralConfiguration;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Windows10General Configuration Request.
 */
public class Windows10GeneralConfigurationRequest extends BaseRequest<Windows10GeneralConfiguration> {
	
    /**
     * The request for the Windows10GeneralConfiguration
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public Windows10GeneralConfigurationRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, Windows10GeneralConfiguration.class);
    }

    /**
     * Gets the Windows10GeneralConfiguration from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super Windows10GeneralConfiguration> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the Windows10GeneralConfiguration from the service
     *
     * @return the Windows10GeneralConfiguration from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Windows10GeneralConfiguration get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super Windows10GeneralConfiguration> callback) {
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
     * Patches this Windows10GeneralConfiguration with a source
     *
     * @param sourceWindows10GeneralConfiguration the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final Windows10GeneralConfiguration sourceWindows10GeneralConfiguration, @Nonnull final ICallback<? super Windows10GeneralConfiguration> callback) {
        send(HttpMethod.PATCH, callback, sourceWindows10GeneralConfiguration);
    }

    /**
     * Patches this Windows10GeneralConfiguration with a source
     *
     * @param sourceWindows10GeneralConfiguration the source object with updates
     * @return the updated Windows10GeneralConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Windows10GeneralConfiguration patch(@Nonnull final Windows10GeneralConfiguration sourceWindows10GeneralConfiguration) throws ClientException {
        return send(HttpMethod.PATCH, sourceWindows10GeneralConfiguration);
    }

    /**
     * Creates a Windows10GeneralConfiguration with a new object
     *
     * @param newWindows10GeneralConfiguration the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final Windows10GeneralConfiguration newWindows10GeneralConfiguration, @Nonnull final ICallback<? super Windows10GeneralConfiguration> callback) {
        send(HttpMethod.POST, callback, newWindows10GeneralConfiguration);
    }

    /**
     * Creates a Windows10GeneralConfiguration with a new object
     *
     * @param newWindows10GeneralConfiguration the new object to create
     * @return the created Windows10GeneralConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Windows10GeneralConfiguration post(@Nonnull final Windows10GeneralConfiguration newWindows10GeneralConfiguration) throws ClientException {
        return send(HttpMethod.POST, newWindows10GeneralConfiguration);
    }

    /**
     * Creates a Windows10GeneralConfiguration with a new object
     *
     * @param newWindows10GeneralConfiguration the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final Windows10GeneralConfiguration newWindows10GeneralConfiguration, @Nonnull final ICallback<? super Windows10GeneralConfiguration> callback) {
        send(HttpMethod.PUT, callback, newWindows10GeneralConfiguration);
    }

    /**
     * Creates a Windows10GeneralConfiguration with a new object
     *
     * @param newWindows10GeneralConfiguration the object to create/update
     * @return the created Windows10GeneralConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Windows10GeneralConfiguration put(@Nonnull final Windows10GeneralConfiguration newWindows10GeneralConfiguration) throws ClientException {
        return send(HttpMethod.PUT, newWindows10GeneralConfiguration);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public Windows10GeneralConfigurationRequest select(@Nonnull final String value) {
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
     public Windows10GeneralConfigurationRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

