#!/usr/bin/perl
use strict;
use warnings;
use utf8;
binmode(STDIN, ':encoding(UTF-8)');
binmode(STDOUT, ':encoding(UTF-8)');

my $infile = $ARGV[0] or die "usage: md2docx.pl input.md output_document.xml\n";
my $outfile = $ARGV[1] or die "usage: md2docx.pl input.md output_document.xml\n";

open(my $in, '<:encoding(UTF-8)', $infile) or die "cannot open $infile: $!";
my @lines = <$in>;
close($in);
chomp(@lines);
# strip any stray CR left over from CRLF line endings
s/\r$// for @lines;

sub xml_escape {
    my ($s) = @_;
    $s =~ s/&/&amp;/g;
    $s =~ s/</&lt;/g;
    $s =~ s/>/&gt;/g;
    return $s;
}

# Convert **bold** and *italic* spans into a sequence of runs. Returns XML string of <w:r>...</w:r> runs.
sub runs_for_text {
    my ($text, $extra_rpr) = @_;
    $extra_rpr = '' unless defined $extra_rpr;
    my @parts = split(/(\*\*.+?\*\*|\*[^*]+?\*)/, $text);
    my $xml = '';
    for my $part (@parts) {
        next if $part eq '';
        if ($part =~ /^\*\*(.+)\*\*$/) {
            my $t = xml_escape($1);
            $xml .= "<w:r><w:rPr><w:b/>$extra_rpr</w:rPr><w:t xml:space=\"preserve\">$t</w:t></w:r>";
        } elsif ($part =~ /^\*([^*]+)\*$/) {
            my $t = xml_escape($1);
            $xml .= "<w:r><w:rPr><w:i/>$extra_rpr</w:rPr><w:t xml:space=\"preserve\">$t</w:t></w:r>";
        } else {
            my $t = xml_escape($part);
            $xml .= "<w:r><w:rPr>$extra_rpr</w:rPr><w:t xml:space=\"preserve\">$t</w:t></w:r>" if $extra_rpr;
            $xml .= "<w:r><w:t xml:space=\"preserve\">$t</w:t></w:r>" unless $extra_rpr;
        }
    }
    return $xml;
}

# Split a markdown table row "| a | b | c |" into trimmed cell strings.
sub split_row {
    my ($line) = @_;
    my $l = $line;
    $l =~ s/^\s*\|//;
    $l =~ s/\|\s*$//;
    my @cells = split(/\|/, $l, -1);
    for my $c (@cells) {
        $c =~ s/^\s+//;
        $c =~ s/\s+$//;
    }
    return @cells;
}

my $body = '';
my $sawTitle = 0;
my $sawSubtitle = 0;
my $pendingLabel = undef;
my $inReferences = 0;
my $tableCounter = 0;
my $figureCounter = 0;
my @chartRows = (); # sidecar lines: id \t type \t title \t labels \t values

my %labelSet = map { $_ => 1 } ('Resumen', 'Abstract', 'Palabras clave', 'Keywords');

# Soft pastel palette shared conceptually with the PowerShell chart renderer.
my $headerFill = 'D6EAF8'; # soft blue
my $altFill    = 'F4F8FB'; # near-white blue tint

for (my $i = 0; $i <= $#lines; $i++) {
    my $line = $lines[$i];
    next if $line =~ /^\s*$/;
    next if $line =~ /^---\s*$/;

    if (!$sawTitle && $line =~ /^#\s+(.+)$/) {
        my $t = xml_escape($1);
        $body .= "<w:p><w:pPr><w:jc w:val=\"center\"/><w:spacing w:after=\"120\"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr><w:t xml:space=\"preserve\">$t</w:t></w:r></w:p>\n";
        $sawTitle = 1;
        next;
    }

    if ($sawTitle && !$sawSubtitle && $line =~ /^##\s+(.+)$/) {
        my $t = xml_escape($1);
        $body .= "<w:p><w:pPr><w:jc w:val=\"center\"/><w:spacing w:after=\"240\"/></w:pPr><w:r><w:rPr><w:b/><w:i/><w:sz w:val=\"26\"/><w:szCs w:val=\"26\"/></w:rPr><w:t xml:space=\"preserve\">$t</w:t></w:r></w:p>\n";
        $sawSubtitle = 1;
        next;
    }

    if ($line =~ /^###\s+(.+)$/) {
        my $t = xml_escape($1);
        $body .= "<w:p><w:pPr><w:pStyle w:val=\"Heading2\"/></w:pPr><w:r><w:t xml:space=\"preserve\">$t</w:t></w:r></w:p>\n";
        next;
    }

    if ($line =~ /^##\s+(.+)$/) {
        my $heading = $1;
        if (exists $labelSet{$heading}) {
            $pendingLabel = $heading;
            next;
        }
        $inReferences = ($heading =~ /^Referencias$/) ? 1 : 0;
        my $t = xml_escape($heading);
        $body .= "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t xml:space=\"preserve\">$t</w:t></w:r></w:p>\n";
        next;
    }

    if ($inReferences) {
        my $runsXml = runs_for_text($line);
        if ($line =~ /^\*Nota:/) {
            $body .= "<w:p><w:pPr><w:jc w:val=\"both\"/></w:pPr>$runsXml</w:p>\n";
        } else {
            $body .= "<w:p><w:pPr><w:ind w:left=\"720\" w:hanging=\"720\"/><w:jc w:val=\"both\"/></w:pPr>$runsXml</w:p>\n";
        }
        next;
    }

    # Table caption directive: [TABLA: texto] immediately before a pipe table.
    if ($line =~ /^\[TABLA:\s*(.+)\]$/) {
        my $caption = $1;
        my $nextLine = '';
        for (my $j = $i + 1; $j <= $#lines; $j++) {
            next if $lines[$j] =~ /^\s*$/;
            $nextLine = $lines[$j];
            last;
        }
        if ($nextLine =~ /^\s*\|/) {
            $tableCounter++;
            my $capEsc = xml_escape($caption);
            $body .= "<w:p><w:pPr><w:jc w:val=\"center\"/><w:spacing w:after=\"80\"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr><w:t xml:space=\"preserve\">Tabla $tableCounter. $capEsc</w:t></w:r></w:p>\n";
            next;
        }
        # not followed by a table; fall through and treat as normal text
    }

    # Markdown pipe table: consume contiguous "| ... |" lines.
    if ($line =~ /^\s*\|/) {
        my @rows = ();
        my $j = $i;
        while ($j <= $#lines && $lines[$j] =~ /^\s*\|/) {
            push @rows, $lines[$j];
            $j++;
        }
        # second row must be the header separator (---|---) to confirm this is a table
        if (@rows >= 2 && $rows[1] =~ /^\s*\|?[\s:\-\|]+\|?\s*$/) {
            my @header = split_row($rows[0]);
            my $ncols = scalar(@header);
            my @bodyRows = ();
            for (my $k = 2; $k < @rows; $k++) {
                push @bodyRows, [ split_row($rows[$k]) ];
            }

            my $usableWidth = 9360; # twips: 12240 letter width - 2*1440 margins
            my $colWidth = $ncols > 0 ? int($usableWidth / $ncols) : $usableWidth;

            my $tbl = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>'
                . '<w:top w:val="single" w:sz="4" w:space="0" w:color="AAB7C4"/>'
                . '<w:left w:val="single" w:sz="4" w:space="0" w:color="AAB7C4"/>'
                . '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="AAB7C4"/>'
                . '<w:right w:val="single" w:sz="4" w:space="0" w:color="AAB7C4"/>'
                . '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="AAB7C4"/>'
                . '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="AAB7C4"/>'
                . '</w:tblBorders><w:tblCellMar>'
                . '<w:top w:w="60" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>'
                . '</w:tblCellMar><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>';
            for (my $c = 0; $c < $ncols; $c++) {
                $tbl .= "<w:gridCol w:w=\"$colWidth\"/>";
            }
            $tbl .= '</w:tblGrid>';

            # header row
            $tbl .= '<w:tr><w:trPr><w:tblHeader/></w:trPr>';
            for my $cell (@header) {
                my $runsXml = runs_for_text($cell, '<w:b/><w:sz w:val="20"/><w:szCs w:val="20"/>');
                $tbl .= "<w:tc><w:tcPr><w:tcW w:w=\"$colWidth\" w:type=\"dxa\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"$headerFill\"/><w:vAlign w:val=\"center\"/></w:tcPr>"
                     . "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>$runsXml</w:p></w:tc>";
            }
            $tbl .= '</w:tr>';

            # body rows, alternating soft shading
            my $rowIdx = 0;
            for my $row (@bodyRows) {
                my $fill = ($rowIdx % 2 == 0) ? $altFill : 'FFFFFF';
                $tbl .= '<w:tr>';
                for (my $c = 0; $c < $ncols; $c++) {
                    my $cellText = defined $row->[$c] ? $row->[$c] : '';
                    my $runsXml = runs_for_text($cellText, '<w:sz w:val="20"/><w:szCs w:val="20"/>');
                    my $jc = ($c == 0) ? 'left' : 'center';
                    $tbl .= "<w:tc><w:tcPr><w:tcW w:w=\"$colWidth\" w:type=\"dxa\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"$fill\"/><w:vAlign w:val=\"center\"/></w:tcPr>"
                         . "<w:p><w:pPr><w:jc w:val=\"$jc\"/></w:pPr>$runsXml</w:p></w:tc>";
                }
                $tbl .= '</w:tr>';
                $rowIdx++;
            }
            $tbl .= '</w:tbl>';

            $body .= "$tbl\n<w:p><w:pPr><w:spacing w:after=\"200\"/></w:pPr></w:p>\n";
            $i = $j - 1; # advance outer loop past the consumed table lines
            next;
        }
        # not a valid table (no separator row) -- fall through to normal paragraph handling
    }

    # Chart directive: fenced ```chart ... ``` block with key: value lines.
    if ($line =~ /^```chart\s*$/) {
        my %meta = (type => 'bar', title => '', labels => '', values => '');
        my $j = $i + 1;
        while ($j <= $#lines && $lines[$j] !~ /^```\s*$/) {
            if ($lines[$j] =~ /^\s*(type|title|labels|values)\s*:\s*(.+?)\s*$/) {
                $meta{$1} = $2;
            }
            $j++;
        }
        $figureCounter++;
        my $chartId = "chart_$figureCounter";
        my $titleEsc = xml_escape($meta{title});
        push @chartRows, join("\t", $chartId, $meta{type}, $meta{title}, $meta{labels}, $meta{values});

        # placeholder paragraph (centered); the build script replaces this run with the actual image.
        $body .= "<w:p><w:pPr><w:jc w:val=\"center\"/><w:spacing w:after=\"80\"/></w:pPr><w:r><w:t xml:space=\"preserve\">##CHART_PLACEHOLDER:$chartId##</w:t></w:r></w:p>\n";
        # caption paragraph
        $body .= "<w:p><w:pPr><w:jc w:val=\"center\"/><w:spacing w:after=\"240\"/></w:pPr><w:r><w:rPr><w:i/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr><w:t xml:space=\"preserve\">Figura $figureCounter. $titleEsc. Datos ilustrativos con fines acad\x{e9}micos.</w:t></w:r></w:p>\n";

        $i = $j; # skip to closing fence
        next;
    }

    # regular paragraph line
    if (defined $pendingLabel) {
        my $labelEsc = xml_escape($pendingLabel);
        my $runsXml = runs_for_text($line);
        $body .= "<w:p><w:pPr><w:jc w:val=\"both\"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\">$labelEsc. </w:t></w:r>$runsXml</w:p>\n";
        $pendingLabel = undef;
        next;
    }

    my $runsXml = runs_for_text($line);
    $body .= "<w:p><w:pPr><w:jc w:val=\"both\"/></w:pPr>$runsXml</w:p>\n";
}

my $sectPr = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

my $document = <<"XML";
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
$body$sectPr
</w:body>
</w:document>
XML

open(my $out, '>:encoding(UTF-8)', $outfile) or die "cannot open $outfile: $!";
print $out $document;
close($out);

if (@chartRows) {
    open(my $ch, '>:encoding(UTF-8)', "$outfile.charts.tsv") or die "cannot open $outfile.charts.tsv: $!";
    print $ch "$_\n" for @chartRows;
    close($ch);
}

print "Generated $outfile with " . scalar(() = $body =~ /<w:p>/g) . " paragraphs, " . scalar(@chartRows) . " charts\n";
