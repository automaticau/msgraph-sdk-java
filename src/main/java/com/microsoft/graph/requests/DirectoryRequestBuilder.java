// Template Source: BaseEntityRequestBuilder.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.Directory;
import com.microsoft.graph.requests.AdministrativeUnitCollectionRequestBuilder;
import com.microsoft.graph.requests.AdministrativeUnitRequestBuilder;
import com.microsoft.graph.requests.DirectoryObjectCollectionRequestBuilder;
import com.microsoft.graph.requests.DirectoryObjectRequestBuilder;
import com.microsoft.graph.requests.UserRequestBuilder;
import com.microsoft.graph.requests.GroupRequestBuilder;
import com.microsoft.graph.requests.ApplicationRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequestBuilder;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Directory Request Builder.
 */
public class DirectoryRequestBuilder extends BaseRequestBuilder<Directory> {

    /**
     * The request builder for the Directory
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public DirectoryRequestBuilder(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }

    /**
     * Creates the request
     *
     * @param requestOptions the options for this request
     * @return the DirectoryRequest instance
     */
    @Nonnull
    public DirectoryRequest buildRequest(@Nullable final com.microsoft.graph.options.Option... requestOptions) {
        return buildRequest(getOptions(requestOptions));
    }

    /**
     * Creates the request with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for this request
     * @return the DirectoryRequest instance
     */
    @Nonnull
    public DirectoryRequest buildRequest(@Nonnull final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        return new com.microsoft.graph.requests.DirectoryRequest(getRequestUrl(), getClient(), requestOptions);
    }


    /**
     *  Gets a request builder for the AdministrativeUnit collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public AdministrativeUnitCollectionRequestBuilder administrativeUnits() {
        return new AdministrativeUnitCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("administrativeUnits"), getClient(), null);
    }

    /**
     * Gets a request builder for the AdministrativeUnit item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public AdministrativeUnitRequestBuilder administrativeUnits(@Nonnull final String id) {
        return new AdministrativeUnitRequestBuilder(getRequestUrlWithAdditionalSegment("administrativeUnits") + "/" + id, getClient(), null);
    }
    /**
     *  Gets a request builder for the DirectoryObject collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public DirectoryObjectCollectionRequestBuilder deletedItems() {
        return new DirectoryObjectCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("deletedItems"), getClient(), null);
    }

    /**
     * Gets a request builder for the DirectoryObject item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public DirectoryObjectRequestBuilder deletedItems(@Nonnull final String id) {
        return new DirectoryObjectRequestBuilder(getRequestUrlWithAdditionalSegment("deletedItems") + "/" + id, getClient(), null);
    }
    /**
     *  Gets a request builder for the User collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public UserCollectionRequestBuilder deletedItemsAsUser() {
        return new UserCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("deletedItems") + "/microsoft.graph.user", getClient(), null);
    }

    /**
     * Gets a request builder for the User item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public UserRequestBuilder deletedItemsAsUser(@Nonnull final String id) {
        return new UserRequestBuilder(getRequestUrlWithAdditionalSegment("deletedItems") + "/" + id + "/microsoft.graph.user", getClient(), null);
    }
    /**
     *  Gets a request builder for the Group collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public GroupCollectionRequestBuilder deletedItemsAsGroup() {
        return new GroupCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("deletedItems") + "/microsoft.graph.group", getClient(), null);
    }

    /**
     * Gets a request builder for the Group item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public GroupRequestBuilder deletedItemsAsGroup(@Nonnull final String id) {
        return new GroupRequestBuilder(getRequestUrlWithAdditionalSegment("deletedItems") + "/" + id + "/microsoft.graph.group", getClient(), null);
    }
    /**
     *  Gets a request builder for the Application collection
     *
     * @return the collection request builder
     */
    @Nonnull
    public ApplicationCollectionRequestBuilder deletedItemsAsApplication() {
        return new ApplicationCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("deletedItems") + "/microsoft.graph.application", getClient(), null);
    }

    /**
     * Gets a request builder for the Application item
     *
     * @return the request builder
     * @param id the item identifier
     */
    @Nonnull
    public ApplicationRequestBuilder deletedItemsAsApplication(@Nonnull final String id) {
        return new ApplicationRequestBuilder(getRequestUrlWithAdditionalSegment("deletedItems") + "/" + id + "/microsoft.graph.application", getClient(), null);
    }
}