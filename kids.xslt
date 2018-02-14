<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" indent="yes" omit-xml-declaration="no"/>
	<xsl:strip-space elements="*"/>
	<xsl:template match="node()|@*">
		<xsl:copy>
			<xsl:apply-templates select="node()|@*"/>
		</xsl:copy>
	</xsl:template>
	<xsl:key name="k" match="event" use="concat(title, '|', RelatedLocations)"/>
	<xsl:template match="events">
		<xsl:copy>
			<xsl:for-each select="event[count(. | key('k', concat(title, '|', RelatedLocations))[1]) = 1]">
				<xsl:sort select="title" />
				<event>
					<xsl:apply-templates select="EventType" />
					<title>
					<xsl:value-of select="title" />
					</title>
					<xsl:for-each select="key('k', concat(title, '|', RelatedLocations))">
						<xsl:sort select="RelatedLocations" />
						<RelatedLocations>
						<xsl:value-of select="RelatedLocations" />
						</RelatedLocations>
					</xsl:for-each>
					<xsl:apply-templates select="Date" />                
					<xsl:apply-templates select="DateYear" />
					<xsl:apply-templates select="DateMonth" />
					<xsl:apply-templates select="DateDay" />
					<DateDay>NA</DateDay>
					<xsl:apply-templates select="Body" />
					<Body></Body>
					<xsl:apply-templates select="AgeRanges" />
					<AgeRanges>NA</AgeRanges>
					<xsl:apply-templates select="RegistrationRequired" />
					<RegistrationRequired></RegistrationRequired>
					<xsl:apply-templates select="RecommendedFor" />
					<RecommendedFor>NA</RecommendedFor>
					<xsl:apply-templates select="Location" />
					<Location>NA</Location>
				</event>                
			</xsl:for-each>
		</xsl:copy>		
	</xsl:template>
	<xsl:template match="*/text()[normalize-space()]">
		<xsl:value-of select="normalize-space()"/>
	</xsl:template>
	<xsl:template match="*/text()[not(normalize-space())]" />
</xsl:stylesheet>