ECG (Etereal Chaos Generator)
By: David Midkiff <mdj2023@hotmail.com>

Description: ECG is a fast, strong, secure and unpredictable pseudo random number generator for Visual Basic.

The Visual Basic "Randomize" command is a function containing a small congruential equation. When Randomize is passed a given number (seed) it returns a number using a linear congruential equation that crunches the seed. Since the Timer in VB is ever-changing most VB programmers use the Timer for the seed. This is a horrible source to use a seed for. The Timer is not random but rather completely predictable. If a crypto salt algorithm uses the Timer as a source for randomness then a potential attacker could attack the system by attacking the PRNG (exhausting timer positions to find patterns). What you need is a combination of random sources to generate a random seed for the Randomize function to crunch. 

ECG does just that using a combination of random sources (mouse positions, window captions global unique identifiers, and window dimensions), hash returns (80-bit HBIT algorithm) and fast timers to generate a genuine random seed for the Randomize function. An attacker would have to guess 10 different random seeds (coming from several random sources) to crack the routine. The ECG design is secure for use in cryptographic routines and is the first open source PRNG for Visual Basic.

---

Why the name?

Ethereal \E*the"re*al\, a.
      Consisting of ether; hence, exceedingly light or airy;
      tenuous; spiritlike; characterized by extreme delicacy, as
      form, manner, thought, etc.

Chaos \Cha"os\ (k[=a]"[o^]s), n.
      a state of extreme confusion and disorder


Ethereal = fast, small and delicate
Chaos = random

Need I say more?