function checkCitations() {
	const doc = DocumentApp.getActiveDocument();
	const body = doc.getBody();

	const paragraphs = body.getParagraphs();
	let text = '';
	paragraphs.forEach( ( p ) => {
		text += p.getText() + '\n';
	} );

	function cleanAuthor( name ) {
		return name
			.trim()
			.replace( /’s$|\'s$/i, '' )
			.replace( /[^\w\-']/g, '' )
			.replace( /\.+$/, '' );
	}

	const citations = {};
	const citationRegex = /\(([^()]+?)\)/g;
	let match;

	while ( ( match = citationRegex.exec( text ) ) !== null ) {
		const fullCitation = match[ 1 ];
		const subCitations = fullCitation.split( /\s*;\s*/ );

		subCitations.forEach( ( sub ) => {
			const inMatch = /([\w\-’']+)\s+in\s+([\w\-’']+)\s+(\d{4}[a-z]?)/i.exec( sub );
			if ( inMatch ) {
				const citedAuthor = cleanAuthor( inMatch[ 1 ] );
				const sourceAuthor = cleanAuthor( inMatch[ 2 ] );
				const year = inMatch[ 3 ];

				if ( citedAuthor.length > 0 ) {
					citations[ citedAuthor + ' ' + year ] = true;
				}
				if ( sourceAuthor.length > 0 ) {
					citations[ sourceAuthor + ' ' + year ] = true;
				}
				return;
			}

			const subMatch = /(.*?)\s+(\d{4}[a-z]?)/.exec( sub );
			if ( subMatch ) {
				const authorsPart = subMatch[ 1 ].trim();
				const year = subMatch[ 2 ];

				if ( /et\.?\s*al\.?/i.test( authorsPart ) ) {
					const firstAuthor = cleanAuthor( authorsPart.split( /\s+/ )[ 0 ] );
					if ( firstAuthor.length > 0 ) {
						citations[ firstAuthor + ' ' + year ] = true;
					}
				} else {
					const authors = authorsPart.split( /\s*(?:,|&|and)\s*/ );
					authors.forEach( ( authorName ) => {
						const author = cleanAuthor( authorName );
						if ( author.length > 0 ) {
							citations[ author + ' ' + year ] = true;
						}
					} );
				}
			}
		} );
	}

	const authorYearRegex = /([\w\-’']+)\s*\((\d{4}[a-z]?)\)/g;
	while ( ( match = authorYearRegex.exec( text ) ) !== null ) {
		const author = cleanAuthor( match[ 1 ] );
		const year = match[ 2 ];
		if ( author.length > 0 && year ) {
			citations[ author + ' ' + year ] = true;
		}
	}

	const bibLines = text.split( /\n|\r/ );
	const bibliography = {};

	bibLines.forEach( ( line ) => {
		line = line.trim();
		if ( line.length === 0 ) return;
		const bibMatch = /^(.+?)\s*\((\d{4}[a-z]?)\)/.exec( line );
		if ( bibMatch ) {
			const authorsPart = bibMatch[ 1 ].trim();
			const year = bibMatch[ 2 ];

			if ( /et\.?\s*al\.?/i.test( authorsPart ) ) {
				const leadAuthorMatch = /^([\w\-’']+)/.exec( authorsPart );
				if ( leadAuthorMatch ) {
					const surname = cleanAuthor( leadAuthorMatch[ 1 ] );
					bibliography[ surname + ' ' + year ] = true;
				}
			} else {
				const authors = authorsPart.split( /\s*(?:,|&|and)\s*/ );
				authors.forEach( ( authorName ) => {
					const author = cleanAuthor( authorName );
					if ( author.length > 0 ) {
						bibliography[ author + ' ' + year ] = true;
					}
				} );
			}
		}
	} );

	const missingCitations = [];
	for ( const c in citations ) {
		if ( ! bibliography.hasOwnProperty( c ) ) {
			missingCitations.push( c );
		}
	}

	if ( missingCitations.length === 0 ) {
		showSidebar( false, '' );
	} else {
		showSidebar( true, missingCitations );
	}
}

function showSidebar( anyIssues, missingEntries ) {
	let html =
		'<style>' +
		'body { font-family: Arial, sans-serif; padding: 10px; }' +
		'.entry { display: block; padding: 8px 0; border-bottom: 1px solid #ccc; }' +
		'.warning { color: #d9534f; font-weight: bold; margin-bottom: 10px; }' +
		'.success { color: #28a745; font-weight: bold; margin-bottom: 10px; }' +
		'</style>';

	if ( anyIssues ) {
		html += '<div class="warning">⚠️ Missing bibliography entries:</div><div>';
		missingEntries.forEach( ( entry ) => {
			html += '<span class="entry">' + entry + '</span>';
		} );
		html += '</div>';
	} else {
		html += '<div class="success">✅ All citations have matching bibliography entries.</div>';
	}

	const htmlOutput = HtmlService.createHtmlOutput( html ).setTitle( 'Missing Citation Checker' );
	DocumentApp.getUi().showSidebar( htmlOutput );
}

function onOpen() {
	DocumentApp.getUi().createAddonMenu().addItem( 'Check citations', 'checkCitations' ).addToUi();
}

function onInstall( e ) {
	onOpen( e );
}
