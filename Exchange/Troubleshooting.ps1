[PS] C:\Windows\system32>Test-FederationTrust -UserIdentity Ashley.Forde@acceptance.mfat.govt.nz -verbose


Begin process.

STEP 1 of 6: Getting ADUser information for Ashley.Forde@acceptance.mfat.govt.nz...
RESULT: Success.

STEP 2 of 6: Getting FederationTrust object for Ashley.Forde@acceptance.mfat.govt.nz...
RESULT: Success.

STEP 3 of 6: Validating that the FederationTrust has the same STS certificates as the actual certificates published by the STS in the federation metadata.
RESULT: Success.

STEP 4 of 6: Getting STS and Organization certificates from the federation trust object...
RESULT: Success.


Validating current configuration for FYDIBOHF25SPDLT.acceptance.mfat.govt.nz...


Validation successful.

STEP 5 of 6: Requesting delegation token...
RESULT: Success. Token retrieved.

Closing Test-FederationTrust...


RunspaceId : d33fe49e-ba3d-4c09-9cd5-58b91c605f22
Id         : FederationTrustConfiguration
Type       : Success
Message    : FederationTrust object in ActiveDirectory is valid.

RunspaceId : d33fe49e-ba3d-4c09-9cd5-58b91c605f22
Id         : FederationMetadata
Type       : Error
Message    : Unable to retrieve federation metadata from the security token service.

RunspaceId : d33fe49e-ba3d-4c09-9cd5-58b91c605f22
Id         : StsCertificate
Type       : Success
Message    : Valid certificate referenced by property TokenIssuerCertificate in the FederationTrust object.

RunspaceId : d33fe49e-ba3d-4c09-9cd5-58b91c605f22
Id         : StsPreviousCertificate
Type       : Success
Message    : Valid certificate referenced by property TokenIssuerPrevCertificate in the FederationTrust object.

RunspaceId : d33fe49e-ba3d-4c09-9cd5-58b91c605f22
Id         : OrganizationCertificate
Type       : Success
Message    : Valid certificate referenced by property OrgPrivCertificate in the FederationTrust object.

RunspaceId : d33fe49e-ba3d-4c09-9cd5-58b91c605f22
Id         : TokenRequest
Type       : Error
Message    : Failed to request delegation token.

Error. Attempted to get delegation token, but token came back as null.
    + CategoryInfo          : NotSpecified: (:) [], LocalizedException
    + FullyQualifiedErrorId : [Server=OA4T3MSGEX01,RequestId=d3866169-48f3-4355-944f-100e74f3c553,TimeStamp=5/9/2022 2:59:39 AM] [FailureCategory=Cmdlet-LocalizedE
   xception] F589C7F4
    + PSComputerName        : oa4t3msgex01.orangetst.mfat.net.nz
