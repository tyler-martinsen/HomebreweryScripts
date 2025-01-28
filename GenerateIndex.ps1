#region ChooseFile
$FilePath = Read-Host "Enter Path to File:"

# Blacklist of regex patterns
$blacklistPatterns = @(
    '^\d+$',                   # Exclude numbers
    '^.*[a-zA-Z]\d.*$'         # Exclude contains number after alpha
    '^\d.*$',                  # Exclude starts with numbers
    '^.*_.*$',                 # Exclude underscore
    '^.*[^\x00-\x7F]+.*$',     # Exclude starts with underscore
    '^http',                   # Exclude URLs
    '^temp',                   # Exclude words starting with 'temp'
    '^[a-zA-Z]{1,2}$',         # Exclude very short words (1-2 letters)
    '^com$',                   # Exclude word com
    '^abilities$',             # Exclude word abilities
    '^ability$',               # Exclude word ability
    '^able$',                  # Exclude word able
    '^absolute$',              # Exclude word absolute
    '^access$',                # Exclude word access
    '^across$',                # Exclude word across
    '^act$',                   # Exclude word act
    '^action$',                # Exclude word action
    '^actions$',               # Exclude word actions
    '^active$',                # Exclude word active
    '^add$',                   # Exclude word add
    '^added$',                 # Exclude word added
    '^additional$',            # Exclude word additional
    '^adjacent$',              # Exclude word adjacent
    '^adjustment$',            # Exclude word adjustment
    '^adulthood$',             # Exclude word adulthood
    '^advanced$',              # Exclude word advanced
    '^affected$',              # Exclude word affected
    '^after$',                 # Exclude word after
    '^again$',                 # Exclude word again
    '^against$',               # Exclude word against
    '^age$',                   # Exclude word age
    '^air$',                   # Exclude word air
    '^all$',                   # Exclude word all
    '^allies$',                # Exclude word allies
    '^allows$',                # Exclude word allows
    '^also$',                  # Exclude word also
    '^ammo$',                  # Exclude word ammo
    '^amount$',                # Exclude word amount
    '^and$',                   # Exclude word and
    '^andrew$',                # Exclude word andrew
    '^another$',               # Exclude word another
    '^any$',                   # Exclude word any
    '^anything$',              # Exclude word anything
    '^appear$',                # Exclude word appear
    '^appearance$',            # Exclude word appearance
    '^appropriate$',           # Exclude word appropriate
    '^are$',                   # Exclude word are
    '^area$',                  # Exclude word area
    '^arm$',                   # Exclude word arm
    '^armor$',                 # Exclude word armor
    '^around$',                # Exclude word around
    '^arrows$',                # Exclude word arrows
    '^artist$',                # Exclude word artist
    '^artstation$',            # Exclude word artstation
    '^assault$',               # Exclude word assault
    '^atmosphere$',            # Exclude word atmosphere
    '^attached$',              # Exclude word attached
    '^attack$',                # Exclude word attack
    '^attacker$',              # Exclude word attacker
    '^attacks$',               # Exclude word attacks
    '^attempt$',               # Exclude word attempt
    '^attribute$',             # Exclude word attribute
    '^authuser$',              # Exclude word authuser
    '^auto$',                  # Exclude word auto
    '^automatically$',         # Exclude word automatically
    '^available$',             # Exclude word available
    '^back$',                  # Exclude word back
    '^background$',            # Exclude word background
    '^balanced$',              # Exclude word balanced
    '^base$',                  # Exclude word base
    '^based$',                 # Exclude word based
    '^because$',               # Exclude word because
    '^become$',                # Exclude word become
    '^been$',                  # Exclude word been
    '^before$',                # Exclude word before
    '^begin$',                 # Exclude word begin
    '^beginning$',             # Exclude word beginning
    '^begins$',                # Exclude word begins
    '^being$',                 # Exclude word being
    '^below$',                 # Exclude word below
    '^benefits$',              # Exclude word benefits
    '^between$',               # Exclude word between
    '^bid$',                   # Exclude word bid
    '^black$',                 # Exclude word black
    '^blend$',                 # Exclude word blend
    '^block$',                 # Exclude word block
    '^blood$',                 # Exclude word blood
    '^body$',                  # Exclude word body
    '^bone$',                  # Exclude word bone
    '^bonus$',                 # Exclude word bonus
    '^bonuses$',               # Exclude word bonuses
    '^botanical$',             # Exclude word botanical
    '^both$',                  # Exclude word both
    '^breaker$',               # Exclude word breaker
    '^build$',                 # Exclude word build
    '^built$',                 # Exclude word built
    '^burst$',                 # Exclude word burst
    '^but$',                   # Exclude word but
    '^called$',                # Exclude word called
    '^can$',                   # Exclude word can
    '^cannot$',                # Exclude word cannot
    '^capable$',               # Exclude word capable
    '^casing$',                # Exclude word casing
    '^cast$',                  # Exclude word cast
    '^cause$',                 # Exclude word cause
    '^chance$',                # Exclude word chance
    '^change$',                # Exclude word change
    '^changes$',               # Exclude word changes
    '^character$',             # Exclude word character
    '^check$',                 # Exclude word check
    '^checks$',                # Exclude word checks
    '^choice$',                # Exclude word choice
    '^choose$',                # Exclude word choose
    '^chosen$',                # Exclude word chosen
    '^claims$',                # Exclude word claims
    '^class$',                 # Exclude word class
    '^classtable$',            # Exclude word classtable
    '^close$',                 # Exclude word close
    '^collection$',            # Exclude word collection
    '^color$',                 # Exclude word color
    '^column$',                # Exclude word column
    '^combat$',                # Exclude word combat
    '^comes$',                 # Exclude word comes
    '^common$',                # Exclude word common
    '^commonwealth$',          # Exclude word commonwealth
    '^competencies$',          # Exclude word competencies
    '^considered$',            # Exclude word considered
    '^control$',               # Exclude word control
    '^copy$',                  # Exclude word copy
    '^core$',                  # Exclude word core
    '^cosmic$',                # Exclude word cosmic
    '^cost$',                  # Exclude word cost
    '^create$',                # Exclude word create
    '^created$',               # Exclude word created
    '^creature$',              # Exclude word creature
    '^creatures$',             # Exclude word creatures
    '^credit$',                # Exclude word credit
    '^credits$',               # Exclude word credits
    '^current$',               # Exclude word current
    '^currently$',             # Exclude word currently
    '^damage$',                # Exclude word damage
    '^damaged$',               # Exclude word damaged
    '^day$',                   # Exclude word day
    '^days$',                  # Exclude word days
    '^deal$',                  # Exclude word deal
    '^dealing$',               # Exclude word dealing
    '^deals$',                 # Exclude word deals
    '^dealt$',                 # Exclude word dealt
    '^decoration$',            # Exclude word decoration
    '^description$',           # Exclude word description
    '^descriptive$',           # Exclude word descriptive
    '^destroyed$',             # Exclude word destroyed
    '^destructive$',           # Exclude word destructive
    '^determine$',             # Exclude word determine
    '^different$',             # Exclude word different
    '^direction$',             # Exclude word direction
    '^directly$',              # Exclude word directly
    '^discovered$',            # Exclude word discovered
    '^display$',               # Exclude word display
    '^does$',                  # Exclude word does
    '^double$',                # Exclude word double
    '^down$',                  # Exclude word down
    '^drawbacks$',             # Exclude word drawbacks
    '^due$',                   # Exclude word due
    '^duration$',              # Exclude word duration
    '^during$',                # Exclude word during
    '^each$',                  # Exclude word each
    '^effect$',                # Exclude word effect
    '^effects$',               # Exclude word effects
    '^either$',                # Exclude word either
    '^electromagnetic$',       # Exclude word electromagnetic
    '^emsp$',                  # Exclude word emsp
    '^end$',                   # Exclude word end
    '^ends$',                  # Exclude word ends
    '^enemies$',               # Exclude word enemies
    '^energy$',                # Exclude word energy
    '^enhancement$',           # Exclude word enhancement
    '^enhancements$',          # Exclude word enhancements
    '^enlistment$',            # Exclude word enlistment
    '^equal$',                 # Exclude word equal
    '^era$',                   # Exclude word era
    '^essence$',               # Exclude word essence
    '^even$',                  # Exclude word even
    '^evenal$',                # Exclude word evenal
    '^every$',                 # Exclude word every
    '^example$',               # Exclude word example
    '^execute$',               # Exclude word execute
    '^eye$',                   # Exclude word eye
    '^eyes$',                  # Exclude word eyes
    '^fail$',                  # Exclude word fail
    '^failure$',               # Exclude word failure
    '^faith$',                 # Exclude word faith
    '^feet$',                  # Exclude word feet
    '^field$',                 # Exclude word field
    '^first$',                 # Exclude word first
    '^fix$',                   # Exclude word fix
    '^flag$',                  # Exclude word flag
    '^fly$',                   # Exclude word fly
    '^followers$',             # Exclude word followers
    '^following$',             # Exclude word following
    '^for$',                   # Exclude word for
    '^force$',                 # Exclude word force
    '^forced$',                # Exclude word forced
    '^form$',                  # Exclude word form
    '^frame$',                 # Exclude word frame
    '^free$',                  # Exclude word free
    '^frequency$',             # Exclude word frequency
    '^from$',                  # Exclude word from
    '^full$',                  # Exclude word full
    '^gain$',                  # Exclude word gain
    '^gained$',                # Exclude word gained
    '^gaining$',               # Exclude word gaining
    '^gains$',                 # Exclude word gains
    '^gave$',                  # Exclude word gave
    '^get$',                   # Exclude word get
    '^give$',                  # Exclude word give
    '^given$',                 # Exclude word given
    '^god$',                   # Exclude word god
    '^googleusercontent$',     # Exclude word googleusercontent
    '^grant$',                 # Exclude word grant
    '^granted$',               # Exclude word granted
    '^grants$',                # Exclude word grants
    '^great$',                 # Exclude word great
    '^greater$',               # Exclude word greater
    '^greenwood$',             # Exclude word greenwood
    '^group$',                 # Exclude word group
    '^guide$',                 # Exclude word guide
    '^gun$',                   # Exclude word gun
    '^hackett$',               # Exclude word hackett
    '^had$',                   # Exclude word had
    '^half$',                  # Exclude word half
    '^handed$',                # Exclude word handed
    '^has$',                   # Exclude word has
    '^have$',                  # Exclude word have
    '^having$',                # Exclude word having
    '^heavy$',                 # Exclude word heavy
    '^height$',                # Exclude word height
    '^hidden$',                # Exclude word hidden
    '^high$',                  # Exclude word high
    '^hit$',                   # Exclude word hit
    '^hits$',                  # Exclude word hits
    '^hold$',                  # Exclude word hold
    '^host$',                  # Exclude word host
    '^hostile$',               # Exclude word hostile
    '^hosts$',                 # Exclude word hosts
    '^hour$',                  # Exclude word hour
    '^hours$',                 # Exclude word hours
    '^how$',                   # Exclude word how
    '^however$',               # Exclude word however
    '^ian$',                   # Exclude word ian
    '^ignore$',                # Exclude word ignore
    '^imgur$',                 # Exclude word imgur
    '^immediately$',           # Exclude word immediately
    '^including$',             # Exclude word including
    '^increase$',              # Exclude word increase
    '^increased$',             # Exclude word increased
    '^increases$',             # Exclude word increases
    '^index$',                 # Exclude word index
    '^information$',           # Exclude word information
    '^inside$',                # Exclude word inside
    '^instead$',               # Exclude word instead
    '^interstellar$',          # Exclude word interstellar
    '^into$',                  # Exclude word into
    '^isotope$',               # Exclude word isotope
    '^item$',                  # Exclude word item
    '^items$',                 # Exclude word items
    '^its$',                   # Exclude word its
    '^joined$',                # Exclude word joined
    '^keep$',                  # Exclude word keep
    '^know$',                  # Exclude word know
    '^known$',                 # Exclude word known
    '^large$',                 # Exclude word large
    '^last$',                  # Exclude word last
    '^lasts$',                 # Exclude word lasts
    '^law$',                   # Exclude word law
    '^lbs$',                   # Exclude word lbs
    '^leave$',                 # Exclude word leave
    '^left$',                  # Exclude word left
    '^less$',                  # Exclude word less
    '^license$',               # Exclude word license
    '^life$',                  # Exclude word life
    '^like$',                  # Exclude word like
    '^line$',                  # Exclude word line
    '^listed$',                # Exclude word listed
    '^live$',                  # Exclude word live
    '^living$',                # Exclude word living
    '^location$',              # Exclude word location
    '^long$',                  # Exclude word long
    '^longer$',                # Exclude word longer
    '^lose$',                  # Exclude word lose
    '^loses$',                 # Exclude word loses
    '^lost$',                  # Exclude word lost
    '^low$',                   # Exclude word low
    '^lower$',                 # Exclude word lower
    '^made$',                  # Exclude word made
    '^magical$',               # Exclude word magical
    '^major$',                 # Exclude word major
    '^make$',                  # Exclude word make
    '^making$',                # Exclude word making
    '^many$',                  # Exclude word many
    '^margin$',                # Exclude word margin
    '^market$',                # Exclude word market
    '^max$',                   # Exclude word max
    '^maximum$',               # Exclude word maximum
    '^may$',                   # Exclude word may
    '^meal$',                  # Exclude word meal
    '^means$',                 # Exclude word means
    '^meat$',                  # Exclude word meat
    '^medium$',                # Exclude word medium
    '^member$',                # Exclude word member
    '^memories$',              # Exclude word memories
    '^metal$',                 # Exclude word metal
    '^mimic$',                 # Exclude word mimic
    '^minimum$',               # Exclude word minimum
    '^minutes$',               # Exclude word minutes
    '^mix$',                   # Exclude word mix
    '^mixed$',                 # Exclude word mixed
    '^mixture$',               # Exclude word mixture
    '^mobile$',                # Exclude word mobile
    '^mode$',                  # Exclude word mode
    '^monster$',               # Exclude word monster
    '^more$',                  # Exclude word more
    '^most$',                  # Exclude word most
    '^move$',                  # Exclude word move
    '^moved$',                 # Exclude word moved
    '^moving$',                # Exclude word moving
    '^much$',                  # Exclude word much
    '^multiple$',              # Exclude word multiple
    '^multiply$',              # Exclude word multiply
    '^must$',                  # Exclude word must
    '^name$',                  # Exclude word name
    '^names$',                 # Exclude word names
    '^nanite$',                # Exclude word nanite
    '^nanites$',               # Exclude word nanites
    '^nbsp$',                  # Exclude word nbsp
    '^need$',                  # Exclude word need
    '^negative$',              # Exclude word negative
    '^new$',                   # Exclude word new
    '^next$',                  # Exclude word next
    '^non$',                   # Exclude word non
    '^none$',                  # Exclude word none
    '^normal$',                # Exclude word normal
    '^normally$',              # Exclude word normally
    '^not$',                   # Exclude word not
    '^note$',                  # Exclude word note
    '^number$',                # Exclude word number
    '^object$',                # Exclude word object
    '^objects$',               # Exclude word objects
    '^off$',                   # Exclude word off
    '^offer$',                 # Exclude word offer
    '^often$',                 # Exclude word often
    '^old$',                   # Exclude word old
    '^once$',                  # Exclude word once
    '^one$',                   # Exclude word one
    '^only$',                  # Exclude word only
    '^opacity$',               # Exclude word opacity
    '^open$',                  # Exclude word open
    '^options$',               # Exclude word options
    '^order$',                 # Exclude word order
    '^origin$',                # Exclude word origin
    '^original$',              # Exclude word original
    '^other$',                 # Exclude word other
    '^others$',                # Exclude word others
    '^out$',                   # Exclude word out
    '^over$',                  # Exclude word over
    '^own$',                   # Exclude word own
    '^page$',                  # Exclude word page
    '^part$',                  # Exclude word part
    '^pass$',                  # Exclude word pass
    '^patron$',                # Exclude word patron
    '^pay$',                   # Exclude word pay
    '^penalties$',             # Exclude word penalties
    '^penalty$',               # Exclude word penalty
    '^per$',                   # Exclude word per
    '^permanent$',             # Exclude word permanent
    '^permanently$',           # Exclude word permanently
    '^picture$',               # Exclude word picture
    '^piercing$',              # Exclude word piercing
    '^place$',                 # Exclude word place
    '^planet$',                # Exclude word planet
    '^plasma$',                # Exclude word plasma
    '^png$',                   # Exclude word png
    '^point$',                 # Exclude word point
    '^points$',                # Exclude word points
    '^portions$',              # Exclude word portions
    '^position$',              # Exclude word position
    '^potential$',             # Exclude word potential
    '^power$',                 # Exclude word power
    '^powerful$',              # Exclude word powerful
    '^prefer$',                # Exclude word prefer
    '^price$',                 # Exclude word price
    '^produce$',               # Exclude word produce
    '^provide$',               # Exclude word provide
    '^provides$',              # Exclude word provides
    '^psychadelic$',           # Exclude word psychadelic
    '^purchase$',              # Exclude word purchase
    '^purchased$',             # Exclude word purchased
    '^purposes$',              # Exclude word purposes
    '^quantum$',               # Exclude word quantum
    '^race$',                  # Exclude word race
    '^radius$',                # Exclude word radius
    '^random$',                # Exclude word random
    '^range$',                 # Exclude word range
    '^ranged$',                # Exclude word ranged
    '^rate$',                  # Exclude word rate
    '^reaches$',               # Exclude word reaches
    '^receive$',               # Exclude word receive
    '^reduce$',                # Exclude word reduce
    '^reduced$',               # Exclude word reduced
    '^reduction$',             # Exclude word reduction
    '^regain$',                # Exclude word regain
    '^related$',               # Exclude word related
    '^religious$',             # Exclude word religious
    '^rely$',                  # Exclude word rely
    '^remain$',                # Exclude word remain
    '^remaining$',             # Exclude word remaining
    '^remains$',               # Exclude word remains
    '^remove$',                # Exclude word remove
    '^removed$',               # Exclude word removed
    '^repeat$',                # Exclude word repeat
    '^reputation$',            # Exclude word reputation
    '^require$',               # Exclude word require
    '^required$',              # Exclude word required
    '^requires$',              # Exclude word requires
    '^reserve$',               # Exclude word reserve
    '^reset$',                 # Exclude word reset
    '^resources$',             # Exclude word resources
    '^restore$',               # Exclude word restore
    '^result$',                # Exclude word result
    '^right$',                 # Exclude word right
    '^ritual$',                # Exclude word ritual
    '^roll$',                  # Exclude word roll
    '^round$',                 # Exclude word round
    '^rounded$',               # Exclude word rounded
    '^rounds$',                # Exclude word rounds
    '^rules$',                 # Exclude word rules
    '^same$',                  # Exclude word same
    '^score$',                 # Exclude word score
    '^section$',               # Exclude word section
    '^see$',                   # Exclude word see
    '^seen$',                  # Exclude word seen
    '^select$',                # Exclude word select
    '^self$',                  # Exclude word self
    '^senior$',                # Exclude word senior
    '^set$',                   # Exclude word set
    '^short$',                 # Exclude word short
    '^shot$',                  # Exclude word shot
    '^sight$',                 # Exclude word sight
    '^single$',                # Exclude word single
    '^size$',                  # Exclude word size
    '^skin$',                  # Exclude word skin
    '^slots$',                 # Exclude word slots
    '^small$',                 # Exclude word small
    '^social$',                # Exclude word social
    '^society$',               # Exclude word society
    '^solid$',                 # Exclude word solid
    '^some$',                  # Exclude word some
    '^something$',             # Exclude word something
    '^source$',                # Exclude word source
    '^space$',                 # Exclude word space
    '^special$',               # Exclude word special
    '^specialization$',        # Exclude word specialization
    '^specific$',              # Exclude word specific
    '^speed$',                 # Exclude word speed
    '^spell$',                 # Exclude word spell
    '^spells$',                # Exclude word spells
    '^spend$',                 # Exclude word spend
    '^spending$',              # Exclude word spending
    '^spent$',                 # Exclude word spent
    '^sphere$',                # Exclude word sphere
    '^square$',                # Exclude word square
    '^squares$',               # Exclude word squares
    '^stack$',                 # Exclude word stack
    '^start$',                 # Exclude word start
    '^stfiddlesticks$',        # Exclude word stfiddlesticks
    '^still$',                 # Exclude word still
    '^stone$',                 # Exclude word stone
    '^store$',                 # Exclude word store
    '^stored$',                # Exclude word stored
    '^succeed$',               # Exclude word succeed
    '^success$',               # Exclude word success
    '^such$',                  # Exclude word such
    '^system$',                # Exclude word system
    '^table$',                 # Exclude word table
    '^take$',                  # Exclude word take
    '^taken$',                 # Exclude word taken
    '^takes$',                 # Exclude word takes
    '^taking$',                # Exclude word taking
    '^tall$',                  # Exclude word tall
    '^target$',                # Exclude word target
    '^targeted$',              # Exclude word targeted
    '^targets$',               # Exclude word targets
    '^technology$',            # Exclude word technology
    '^tend$',                  # Exclude word tend
    '^tendency$',              # Exclude word tendency
    '^terminal$',              # Exclude word terminal
    '^than$',                  # Exclude word than
    '^that$',                  # Exclude word that
    '^the$',                   # Exclude word the
    '^their$',                 # Exclude word their
    '^them$',                  # Exclude word them
    '^themselves$',            # Exclude word themselves
    '^then$',                  # Exclude word then
    '^there$',                 # Exclude word there
    '^these$',                 # Exclude word these
    '^they$',                  # Exclude word they
    '^this$',                  # Exclude word this
    '^thomas$',                # Exclude word thomas
    '^those$',                 # Exclude word those
    '^though$',                # Exclude word though
    '^threads$',               # Exclude word threads
    '^through$',               # Exclude word through
    '^time$',                  # Exclude word time
    '^times$',                 # Exclude word times
    '^top$',                   # Exclude word top
    '^total$',                 # Exclude word total
    '^toxin$',                 # Exclude word toxin
    '^trait$',                 # Exclude word trait
    '^transform$',             # Exclude word transform
    '^treated$',               # Exclude word treated
    '^turn$',                  # Exclude word turn
    '^two$',                   # Exclude word two
    '^type$',                  # Exclude word type
    '^types$',                 # Exclude word types
    '^typically$',             # Exclude word typically
    '^under$',                 # Exclude word under
    '^unique$',                # Exclude word unique
    '^unless$',                # Exclude word unless
    '^until$',                 # Exclude word until
    '^upon$',                  # Exclude word upon
    '^use$',                   # Exclude word use
    '^used$',                  # Exclude word used
    '^uses$',                  # Exclude word uses
    '^using$',                 # Exclude word using
    '^value$',                 # Exclude word value
    '^vegetable$',             # Exclude word vegetable
    '^very$',                  # Exclude word very
    '^want$',                  # Exclude word want
    '^was$',                   # Exclude word was
    '^way$',                   # Exclude word way
    '^week$',                  # Exclude word week
    '^weigh$',                 # Exclude word weigh
    '^weight$',                # Exclude word weight
    '^were$',                  # Exclude word were
    '^what$',                  # Exclude word what
    '^when$',                  # Exclude word when
    '^whenever$',              # Exclude word whenever
    '^where$',                 # Exclude word where
    '^which$',                 # Exclude word which
    '^while$',                 # Exclude word while
    '^white$',                 # Exclude word white
    '^who$',                   # Exclude word who
    '^wide$',                  # Exclude word wide
    '^width$',                 # Exclude word width
    '^will$',                  # Exclude word will
    '^willing$',               # Exclude word willing
    '^with$',                  # Exclude word with
    '^within$',                # Exclude word within
    '^without$',               # Exclude word without
    '^work$',                  # Exclude word work
    '^worth$',                 # Exclude word worth
    '^would$',                 # Exclude word would
    '^www$',                   # Exclude word www
    '^year$',                  # Exclude word year
    '^years$',                 # Exclude word years
    '^yes$',                   # Exclude word yes
    '^you$',                   # Exclude word you
    '^your$',                  # Exclude word your
    '^yourself$',              # Exclude word yourself
    '^zero$',                  # Exclude word zero
    '^com$'                    # Exclude word com
)

# Check if the file exists
if (-not (Test-Path -Path $FilePath)) {
    Write-Host "File not found: $FilePath" -ForegroundColor Red
    exit
}

# Read the content of the file
$content = Get-Content -Path $FilePath

# Initialize variables
$pageNumber = 1
$index = @{}

# Function to clean and split text into words
function Get-Words($text) {
    $text -replace '<[^>]*>', '' -split '\W+' | Where-Object { $_ -ne '' }
}

# Check if a word matches any blacklist pattern
function Is-Blacklisted($word) {
    foreach ($pattern in $blacklistPatterns) {
        if ($word -match $pattern) {
            return $true
        }
    }
    return $false
}

# Process each line in the file
foreach ($line in $content) {
    # Detect page number div
    if ($line -match "<div class='pageNumber'>(\d+)</div>") {
        $pageNumber = [int]$Matches[1]
        continue
    }

    if ($line -match "^# Index") {
        $pageNumber = [int]$Matches[1]
        exit
    }

    # Extract words not inside < > tags
    $words = Get-Words $line
    foreach ($word in $words) {
        # Convert word to lowercase for uniformity
        $word = $word.ToLower()

        # Skip blacklisted words
        if (Is-Blacklisted $word) {
            continue
        }

        # Add the word to the index
        if (-not $index.ContainsKey($word)) {
            $index[$word] = @()
        }
        if (-not ($index[$word] -contains $pageNumber)) {
            $index[$word] += $pageNumber
        }
    }
}

# Output the index
$outputPath = Join-Path -Path (Split-Path -Parent $FilePath) -ChildPath "DocumentIndex.txt"
"Word Index for File: $FilePath" | Out-File -FilePath $outputPath -Encoding UTF8

Set-Content -Path $outputPath -Value ""

$Start = "`r`n"
$Start += "# Index`r`n`r`n"
$Start += "{{monster,wide,frame`r`n"

$Mid = "}}`r`n"
$Mid += "`r`n"
$Mid += "<div class='pageNumber'>1</div>`r`n"
$Mid += "<div class='footnote' style='background: rgba(0,0,255,0.2);text-align:center;font-weight:bold;height:20px;line-height:20px;font-size: 12px;width:656px;text-indent: 10px;'>Title</div>`r`n"
$Mid += "\page`r`n"
$Mid += "`r`n"
$Mid += "{{monster,wide,frame`r`n"
$Mid += "`r`n"

$End = "}}`r`n"
$End += "`r`n"
$End += "<div class='pageNumber'>1</div>`r`n"
$End += "<div class='footnote' style='background: rgba(0,0,255,0.2);text-align:center;font-weight:bold;height:20px;line-height:20px;font-size: 12px;width:656px;text-indent: 10px;'>Title</div>`r`n"
$End += "\page`r`n"

$outputBuffer = "$Start`n"  # Initialize buffer with Start

$indexNumber = 0

foreach ($key in $index.Keys | Sort-Object) {

    if ($indexNumber -gt 110) {
        $outputBuffer += "$Mid`n"  # Append `$Mid` to the buffer
        $indexNumber = 0
    } else {
        $indexNumber++
    }

    if ($index[$key].Count -gt 10) {
        Write-Host $key -ForegroundColor Green
    }

    $pages = ($index[$key] | Sort-Object) -join ", "
    $outputBuffer += "- ${key}: $pages`n"  # Append the current key and pages to the buffer
}

$outputBuffer += "$End`n"  # Append `$End` to the buffer

# Write the accumulated output to the file
$outputBuffer | Out-File -FilePath $outputPath -Encoding UTF8

Write-Host "Index created successfully: $outputPath" -ForegroundColor Green
