import os
import sys
import string
import math
from Registry import Registry

#b24chrs = (string.digits + string.ascii_uppercase)[:24]
generic_b24chrs = '0123456789ABCDEFGHIJKLMN'

code_len = 25 # encoded key length (user-readable key)
bin_len = 15 # binary key length
regkey_idx = 52 # start of key in DPID for 2003, 2007
regkey_idx_2010 = 0x328 # start in DPID for 2010

b24chrs = 'BCDFGHJKMPQRTVWXY2346789'
reg_root = r'Software\Microsoft\Office'

def chunks(l, n):
    """ Yield successive n-sized chunks from l.
    """
    for i in xrange(0, len(l), n):
        yield l[i:i+n]
        
def b24encode(input, outlen=None, chrmap=None):

    # The number of coded characters to actually generate can't be 
    # determined soley from the input length, but we can guess.
    if outlen is None:
        # each base24 code char takes ~4.585 bits (ln24 / ln2)
        outlen = int(math.ceil(8*len(input) / 4.585))
        
    # Use default character mapping [0-9A-N] if none provided
    if chrmap is None:
        chrmap = generic_b24chrs
        
    input = [ord(i) for i in input[::-1]]
    '''
    # takes less memory (does it piecewise), but more complex
    decoded = []
    for i in range(0,encoded_chars + 1)[::-1]:
        r = 0
        for j in range(0,15)[::-1]:
            r = (r * 256) ^ input[j]
            input[j] = r / 24
            r = r % 24
        
        print b24chrs[r]
        decoded = decoded.append(b24chrs[r])
    
    return decoded[::-1]
    '''
    
    # simple, but eats a ton of memory and probably time if the 
    # encoded string is large
    enc = 0
    for i in input:
        enc = enc * 256 + i
        
    dec = []
    for i in range(outlen):
        dec.append(chrmap[enc % 24])
        enc = enc // 24
        
    dec.reverse()
    return ''.join(dec)

def b24decode(input, chrmap=None):

    # Use default character mapping [0-9A-N] if none provided
    if chrmap is None:
        chrmap = generic_b24chrs
    
    # clean invalid characters from input (e.g. '-' (dashes) in product key)
    # and map to \x00 through \x23.
    rmchrs = []
    for i in xrange(256):
        if not chr(i) in chrmap:
            rmchrs.append(chr(i))
    tt = string.maketrans(chrmap, ''.join([chr(i) for i in xrange(24)]))
    input = input.translate(tt, ''.join(rmchrs))
        
    encnum = 0
    for cc in input:
        encnum *= 24
        encnum += ord(cc)
        
    enc = []
    while encnum:
        enc.append(encnum % 256)
        encnum = encnum // 256
    
    return ''.join([chr(i) for i in enc])
    
def msoKeyDecode(regkey, version=None):
    '''Decodes a registry key value, by extracting product key 
    from bytes 52-66 and decoding.
    
    Office 2010 (14.0) appears to store the key at 0x328 to 0x337 in 
    DPID.  The "Registration" tree is different (cluttered) versus 
    other versions, and the DPID value is (exactly) 7 times longer than 
    before (1148 bytes, up from 164).
    
    Tested with a 2010 full suite and a trial of Visio.
    
    Parameters:
    - regkey is a string containing the contents of "DigitalProductID"
    - version is the decimal version number given by the key directly 
    under the "Office" key root
    '''
    if version is None:
        version = 11 # (default 2003, 2007 appears to be compatible.)
        
    if float(version) < 14:
        enckey = regkey[regkey_idx:regkey_idx+bin_len]
    else:
        enckey = regkey[regkey_idx_2010:regkey_idx_2010+bin_len]
    
    deckey = b24encode(enckey, code_len, chrmap=b24chrs)
        
    return '-'.join(list(chunks(deckey,5)))
    
def iter_keyvalues(key):
    for kv in key.values():
        yield (kv.name(), kv.value())
    
def iter_all(key):
    for kv in key.values():
        yield (kv.name(), kv.value())
    for subkey in key.subkeys():
        yield (subkey.name(), '')
        for k,v in iter_all(subkey):
            yield (k,v)

def main(argv=None):
    '''Scans local Microsoft Office registry keys for DigitalProductID values
    and encodes the binary data in base24.
    
    Note: The given "Name:" of Office 2010 products is incorrect
    (may just provide a single program name), though the Product Key 
    should be valid.
    '''
    if argv is None:
        argv = sys.argv
        
    if len(argv) != 2 or not os.path.isfile(argv[1]):
        print """Usage:
            %s <SOFTWARE registry file path>
            Search for Microsoft Office product keys in the SOFTWARE registry
            A typical path would be /mnt/image/Windows/system32/config/software
            or c:\\Windows\\system32\\config\\software
        """ % argv[0]
        exit(-1)

    software_registry_filename = argv[1]
    print "Reading registry file", software_registry_filename 

    f = open(software_registry_filename, 'rb')
    r = Registry.Registry(f)
    microsoft = r.root().find_key('Microsoft')
    office = microsoft.find_key('Office')

    prod_keys = []
    
    for subkey1 in office.subkeys():
        # Microsoft\Office subkeys tend to be version numbers (11.0, 12.0, 14.0...)
        for subkey2 in subkey1.subkeys():
            # versions contain Registration
            for subkey3 in subkey2.subkeys():
                # Registration contains DPID
                dpid_found = False
                for k,v in iter_keyvalues(subkey3):
                    if k == 'DigitalProductID':
                        dpid_found = True
                        dpid = v
                    if k == 'ProductName':
                        name = v
                if dpid_found:
                    #print "Product Name: %s, Key: %s, Path: %s" % (name, msoKeyDecode(dpid, subkey1.name()), subkey3.path())
                    prod_keys.append((name, msoKeyDecode(dpid, subkey1.name()), subkey3.path()))

    rf = "{0:<%i} {1:<}" % (max([len(i[0]) for i in prod_keys]) + 3)
    
    product_head = "Product Name"
    dpid_head = "Digital Product ID (key encoded in base24)"
    
    print rf.format(product_head, dpid_head)
    print rf.format('-' * len(product_head),'-' * len(dpid_head))
    
    for prod_key in prod_keys:
        print rf.format(*prod_key)
    
        
if __name__ == "__main__":
    sys.exit(main())
