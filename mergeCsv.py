#!/usr/bin/env python
import argparse, csv
if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='merge csv files on field', version='%(prog)s 1.0')
    parser.add_argument('infile', nargs='+', type=str, help='list of input files')
    parser.add_argument('--out', type=str, default='temp.csv', help='name of output file')
    args = parser.parse_args()
    data = {}
    fields = []

    for fname in args.infile:
        with open(fname) as df:
            reader = csv.DictReader(df)
            for line in reader:
                # assuming the field is called ID
                if line['ID'] not in data:
                    data[line['ID']] = line
                else:
                    for k,v in line.iteritems():
                        if k not in data[line['ID']]:
                            data[line['ID']][k] = v
                for k in line.iterkeys():
                    if k not in fields:
                        fields.append(k)
            del reader

    writer = csv.DictWriter(open(args.out, "wb"), fields, dialect='excel')
    # write the header at the top of the file
    writer.writeheader()
    writer.writerows(data)
    del writer
