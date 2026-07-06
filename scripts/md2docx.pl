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

my $body = '';
my $sawTitle = 0;
my $sawSubtitle = 0;
my $pendingLabel = undef;
my $inReferences = 0;

my %labelSet = map { $_ => 1 } ('Resumen', 'Abstract', 'Palabras clave', 'Keywords');

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

print "Generated $outfile with " . scalar(() = $body =~ /<w:p>/g) . " paragraphs\n";
