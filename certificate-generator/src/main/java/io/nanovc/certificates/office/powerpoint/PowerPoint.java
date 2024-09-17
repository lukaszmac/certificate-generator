package io.nanovc.certificates.office.powerpoint;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Variant;

import java.io.Closeable;
import java.nio.file.Path;

/**
 * This is used for automating PowerPoint.
 * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application">Application Object</a>
 */
public class PowerPoint implements AutoCloseable
{
    /**
     * The PowerPoint application ActiveX Component.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application">Application Object</a>
     */
    protected ActiveXComponent application;

    /**
     * The Presentations ActiveX Component.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.presentations">Presentations Property</a>
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations">Presentations Collection</a>
     */
    protected ActiveXComponent presentations;

    /**
     * Starts the PowerPoint application.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application">Application Object</a>
     */
    public void start()
    {
        // Initialise the COM thread:
        ComThread.InitSTA();

        // Start the application:
        this.application = new ActiveXComponent("PowerPoint.Application");

        // Find components of interest:
        this.presentations = this.application.getPropertyAsComponent("Presentations");
    }

    /**
     * Quits the PowerPoint application.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.quit">Quit Method</a>
     */
    public void quit()
    {
        // Check whether we need to quit:
        if (this.application != null)
        {
            // Quit the application:
            this.application.invoke("Quit");

            // Close the COM thread:
            ComThread.Release();

            // Free the components:
            this.application = null;
            this.presentations = null;
        }
    }

    /**
     * Opens a PowerPoint presentation at the given path.
     * @see <a href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open">Open Method</a>
     * @param pathToPresentation The path to the presentation that we want to open.
     * @return The presentation that was opened.
     */
    public Presentation openPresentation(Path pathToPresentation)
    {
        // Get the absolute path to the presentation:
        String fullPathToPresentation = pathToPresentation.toAbsolutePath().normalize().toString();

        // Open the presentation:
        var presentationObject = this.presentations.invokeGetComponent("Open", new Variant(fullPathToPresentation), new Variant(0), new Variant(0), new Variant(0));

        // Create the presentation wrapper:
        Presentation presentation = new Presentation(presentationObject);

        return presentation;
    }

    /**
     * Closes this resource, relinquishing any underlying resources.
     * This method is invoked automatically on objects managed by the
     * {@code try}-with-resources statement.
     *
     * @throws Exception if this resource cannot be closed
     * @apiNote While this interface method is declared to throw {@code
     *     Exception}, implementers are <em>strongly</em> encouraged to
     *     declare concrete implementations of the {@code close} method to
     *     throw more specific exceptions, or to throw no exception at all
     *     if the close operation cannot fail.
     *
     *     <p> Cases where the close operation may fail require careful
     *     attention by implementers. It is strongly advised to relinquish
     *     the underlying resources and to internally <em>mark</em> the
     *     resource as closed, prior to throwing the exception. The {@code
     *     close} method is unlikely to be invoked more than once and so
     *     this ensures that the resources are released in a timely manner.
     *     Furthermore it reduces problems that could arise when the resource
     *     wraps, or is wrapped, by another resource.
     *
     *     <p><em>Implementers of this interface are also strongly advised
     *     to not have the {@code close} method throw {@link
     *     InterruptedException}.</em>
     *     <p>
     *     This exception interacts with a thread's interrupted status,
     *     and runtime misbehavior is likely to occur if an {@code
     *     InterruptedException} is {@linkplain Throwable#addSuppressed
     *     suppressed}.
     *     <p>
     *     More generally, if it would cause problems for an
     *     exception to be suppressed, the {@code AutoCloseable.close}
     *     method should not throw it.
     *
     *     <p>Note that unlike the {@link Closeable#close close}
     *     method of {@link Closeable}, this {@code close} method
     *     is <em>not</em> required to be idempotent.  In other words,
     *     calling this {@code close} method more than once may have some
     *     visible side effect, unlike {@code Closeable.close} which is
     *     required to have no effect if called more than once.
     *     <p>
     *     However, implementers of this interface are strongly encouraged
     *     to make their {@code close} methods idempotent.
     */
    @Override
    public void close() throws Exception
    {
        quit();
    }
}
